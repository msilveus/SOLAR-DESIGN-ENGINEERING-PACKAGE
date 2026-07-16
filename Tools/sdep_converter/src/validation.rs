use std::collections::BTreeMap;

use regex::Regex;
use serde_json::{Map, Value};

use crate::{parse_json_bool, InputData, SchemaData, ValidationReport};

pub fn validate_database(
    schema: &SchemaData,
    input_data: &InputData,
    verbosity: u8,
) -> Result<ValidationReport, String> {
    let database_name = schema
        .get("database")
        .or_else(|| schema.get("name"))
        .and_then(Value::as_str)
        .ok_or_else(|| "Schema is missing a top-level `database` string".to_string())?;

    match normalize_database_name(database_name).as_str() {
        "panel" | "panels" => {
            let validator = PanelValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "inverter" | "inverters" => {
            let validator = InverterValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "battery" | "batteries" => {
            let validator = BatteryValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "optimizer" | "optimizers" => {
            let validator = BatteryValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "battery disconnect switch" | "battery disconnect switches" => {
            let validator = BatteryValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "pv disconnect switch" | "pv disconnect switches" => {
            let validator = BatteryValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        "rsd" => {
            let validator = BatteryValidator::new(schema, input_data, verbosity);
            validator.validate()
        }
        other => Err(format!(
            "Unsupported database type `{other}` in schema"
        )),
    }
}

trait DatabaseValidator {
    fn validate(&self) -> Result<ValidationReport, String>;
}

struct PanelValidator<'a> {
    schema: &'a SchemaData,
    input_data: &'a InputData,
    verbosity: u8,
}

struct InverterValidator<'a> {
    schema: &'a SchemaData,
    input_data: &'a InputData,
    verbosity: u8,
}

struct BatteryValidator<'a> {
    schema: &'a SchemaData,
    input_data: &'a InputData,
    verbosity: u8,
}

#[derive(Copy, Clone)]
enum ValidationOutcome {
    Pass,
    Fail,
    Warn,
}

impl<'a> PanelValidator<'a> {
    fn new(schema: &'a SchemaData, input_data: &'a InputData, verbosity: u8) -> Self {
        Self {
            schema,
            input_data,
            verbosity,
        }
    }
}

impl<'a> InverterValidator<'a> {
    fn new(schema: &'a SchemaData, input_data: &'a InputData, verbosity: u8) -> Self {
        Self {
            schema,
            input_data,
            verbosity,
        }
    }
}

impl<'a> BatteryValidator<'a> {
    fn new(schema: &'a SchemaData, input_data: &'a InputData, verbosity: u8) -> Self {
        Self {
            schema,
            input_data,
            verbosity,
        }
    }
}

impl<'a> DatabaseValidator for PanelValidator<'a> {
    fn validate(&self) -> Result<ValidationReport, String> {
        validate_rows_with_policy(
            self.schema,
            self.input_data,
            self.verbosity,
            ValidationPolicy {
                apply_panel_rules: true,
                apply_schema_checks: true,
                warn_on_duplicate_name_model: true,
            },
        )
    }
}

impl<'a> DatabaseValidator for InverterValidator<'a> {
    fn validate(&self) -> Result<ValidationReport, String> {
        validate_rows_with_policy(
            self.schema,
            self.input_data,
            self.verbosity,
            ValidationPolicy {
                apply_panel_rules: false,
                apply_schema_checks: true,
                warn_on_duplicate_name_model: false,
            },
        )
    }
}

impl<'a> DatabaseValidator for BatteryValidator<'a> {
    fn validate(&self) -> Result<ValidationReport, String> {
        validate_rows_with_policy(
            self.schema,
            self.input_data,
            self.verbosity,
            ValidationPolicy {
                apply_panel_rules: false,
                apply_schema_checks: true,
                warn_on_duplicate_name_model: false,
            },
        )
    }
}

#[derive(Copy, Clone)]
struct ValidationPolicy {
    apply_panel_rules: bool,
    apply_schema_checks: bool,
    warn_on_duplicate_name_model: bool,
}

fn validate_rows_with_policy(
    schema: &SchemaData,
    input_data: &InputData,
    verbosity: u8,
    policy: ValidationPolicy,
) -> Result<ValidationReport, String> {
    let field_definitions = schema
        .get("fields")
        .and_then(Value::as_object)
        .ok_or_else(|| "Schema is missing a top-level `fields` object".to_string())?;

    let required_fields = schema
        .get("requiredFields")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();

    let primary_key = schema
        .get("primaryKey")
        .and_then(Value::as_str)
        .unwrap_or("ID");

    let mut validation: ValidationReport = BTreeMap::new();
    let mut failed_records = 0usize;
    let mut warnings = Vec::new();
    let mut seen_primary_keys: BTreeMap<String, usize> = BTreeMap::new();
    let mut seen_name_model_pairs: BTreeMap<(String, String), usize> = BTreeMap::new();

    for (row_index, row) in input_data.iter().enumerate() {
        let row_number = row_index + 1;
        let record_label = validation_record_label(row, row_number);
        let mut record_errors = Vec::new();

        for (field_name, field_definition) in field_definitions {
            validate_field(
                field_name,
                field_definition,
                row.get(field_name),
                row_number,
                &record_label,
                &required_fields,
                verbosity,
                &mut record_errors,
            )?;
        }

        if policy.apply_panel_rules {
            validate_panel_engineering(
                row,
                row_number,
                &record_label,
                verbosity,
                &mut record_errors,
            )?;
        }

        if policy.apply_schema_checks {
            validate_engineering_checks(
                schema,
                row,
                row_number,
                &record_label,
                verbosity,
                &mut record_errors,
                &mut warnings,
            )?;
        }

        validate_duplicate_primary_key(
            row,
            row_number,
            &record_label,
            verbosity,
            primary_key,
            &mut record_errors,
            &mut seen_primary_keys,
        );

        if policy.warn_on_duplicate_name_model {
            validate_duplicate_panel_pair(
                row,
                row_number,
                &record_label,
                verbosity,
                &mut warnings,
                &mut seen_name_model_pairs,
            );
        }

        if !record_errors.is_empty() {
            failed_records += 1;
            validation.insert(
                format!("record_{row_number}"),
                Value::Object(Map::from_iter([
                    (
                        "recordNumber".to_string(),
                        Value::from(row_number as u64),
                    ),
                    (
                        "errors".to_string(),
                        Value::Array(record_errors.into_iter().map(Value::String).collect()),
                    ),
                ])),
            );
            if verbosity >= 4 {
                if let Some(created) = validation.get(&format!("record_{row_number}")) {
                    println!("Created validation dictionary for row {row_number}: {created:#?}");
                }
            }
        }
    }

    let passed = failed_records == 0;
    validation.insert("passed".to_string(), Value::Bool(passed));

    if passed {
        validation.insert(
            "validatedRecords".to_string(),
            Value::from(input_data.len() as u64),
        );
    }

    if !warnings.is_empty() {
        validation.insert("warnings".to_string(), Value::Array(warnings));
    }

    Ok(validation)
}

fn validate_engineering_checks(
    schema: &SchemaData,
    row: &Map<String, Value>,
    row_number: usize,
    record_label: &str,
    verbosity: u8,
    record_errors: &mut Vec<String>,
    warnings: &mut Vec<Value>,
) -> Result<(), String> {
    let Some(checks) = schema.get("engineeringChecks").and_then(Value::as_array) else {
        return Ok(());
    };

    for check in checks {
        let check_object = check.as_object().ok_or_else(|| {
            "Each engineering check must be a JSON object".to_string()
        })?;
        let name = check_object
            .get("name")
            .and_then(Value::as_str)
            .unwrap_or("engineering check");
        let severity = check_object
            .get("severity")
            .and_then(Value::as_str)
            .unwrap_or("error");
        let optional = check_object
            .get("optional")
            .and_then(Value::as_bool)
            .unwrap_or(false);

        let result = match check_object.get("type").and_then(Value::as_str) {
            Some("greater_than") => {
                evaluate_numeric_threshold(check_object, row, |left, right| left > right)?
            }
            Some("greater_equal") => {
                evaluate_numeric_threshold(check_object, row, |left, right| left >= right)?
            }
            Some("field_greater_than") => {
                evaluate_field_comparison(check_object, row, |left, right| left > right)?
            }
            Some("field_greater_equal") => {
                evaluate_field_comparison(check_object, row, |left, right| left >= right)?
            }
            Some("integer") => evaluate_integer_check(check_object, row)?,
            Some(other) => {
                return Err(format!("Unsupported engineering check type `{other}`"));
            }
            None => {
                return Err("Engineering check is missing a `type` value".to_string());
            }
        };

        match result {
            Some(true) => log_validation_result(
                verbosity,
                row_number,
                record_label,
                name,
                ValidationOutcome::Pass,
            ),
            Some(false) => {
                log_validation_result(
                    verbosity,
                    row_number,
                    record_label,
                    name,
                    if severity == "warning" {
                        ValidationOutcome::Warn
                    } else {
                        ValidationOutcome::Fail
                    },
                );
                if severity == "warning" {
                    warnings.push(Value::Object(Map::from_iter([
                        ("recordNumber".to_string(), Value::from(row_number as u64)),
                        ("severity".to_string(), Value::String("warning".to_string())),
                        (
                            "message".to_string(),
                            Value::String(format!("Row {row_number}: `{name}` failed")),
                        ),
                    ])));
                } else {
                    record_errors.push(format!("Row {row_number}: `{name}` failed"));
                }
            }
            None if optional => {}
            None => {
                if severity == "warning" {
                    log_validation_result(
                        verbosity,
                        row_number,
                        record_label,
                        name,
                        ValidationOutcome::Warn,
                    );
                } else {
                    log_validation_result(
                        verbosity,
                        row_number,
                        record_label,
                        name,
                        ValidationOutcome::Fail,
                    );
                    record_errors.push(format!("Row {row_number}: `{name}` could not be evaluated"));
                }
            }
        }
    }

    Ok(())
}

fn evaluate_numeric_threshold<F>(
    check_object: &Map<String, Value>,
    row: &Map<String, Value>,
    predicate: F,
) -> Result<Option<bool>, String>
where
    F: Fn(f64, f64) -> bool,
{
    let field_name = check_object
        .get("field")
        .and_then(Value::as_str)
        .ok_or_else(|| "greater_than checks require a `field` string".to_string())?;
    let threshold = check_object
        .get("value")
        .and_then(Value::as_f64)
        .ok_or_else(|| "greater_than checks require a numeric `value`".to_string())?;
    Ok(parse_row_number(row, field_name).map(|value| predicate(value, threshold)))
}

fn evaluate_field_comparison<F>(
    check_object: &Map<String, Value>,
    row: &Map<String, Value>,
    predicate: F,
) -> Result<Option<bool>, String>
where
    F: Fn(f64, f64) -> bool,
{
    let left_field = check_object
        .get("left")
        .and_then(Value::as_str)
        .ok_or_else(|| "field comparison checks require a `left` string".to_string())?;
    let right_field = check_object
        .get("right")
        .and_then(Value::as_str)
        .ok_or_else(|| "field comparison checks require a `right` string".to_string())?;
    let Some(left_value) = parse_row_number(row, left_field) else {
        return Ok(None);
    };
    let Some(right_value) = parse_row_number(row, right_field) else {
        return Ok(None);
    };

    Ok(Some(predicate(left_value, right_value)))
}

fn evaluate_integer_check(
    check_object: &Map<String, Value>,
    row: &Map<String, Value>,
) -> Result<Option<bool>, String> {
    let field_name = check_object
        .get("field")
        .and_then(Value::as_str)
        .ok_or_else(|| "integer checks require a `field` string".to_string())?;
    let Some(raw_value) = row.get(field_name).and_then(Value::as_str).map(str::trim) else {
        return Ok(None);
    };
    if raw_value.is_empty() {
        return Ok(None);
    }

    Ok(Some(raw_value.parse::<i64>().is_ok()))
}

fn validate_field(
    field_name: &str,
    field_definition: &Value,
    field_value: Option<&Value>,
    row_number: usize,
    record_label: &str,
    required_fields: &[Value],
    verbosity: u8,
    record_errors: &mut Vec<String>,
) -> Result<(), String> {
    let field_object = field_definition
        .as_object()
        .ok_or_else(|| format!("Field definition for `{field_name}` must be a JSON object"))?;

    let is_optional = field_object
        .get("optional")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let is_required = if is_optional {
        false
    } else {
        field_object
            .get("required")
            .and_then(Value::as_bool)
            .unwrap_or(false)
            || required_fields
                .iter()
                .any(|value| value.as_str() == Some(field_name))
    };

    let raw_value = field_value.and_then(Value::as_str).unwrap_or("").trim();

    if raw_value.is_empty() {
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            &format!("{field_name} required"),
            if is_required {
                ValidationOutcome::Fail
            } else {
                ValidationOutcome::Pass
            },
        );
        if is_required {
            record_errors.push(format!("Row {row_number}: `{field_name}` is required"));
        }
        return Ok(());
    }

    let field_type = field_object
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            format!("Field definition for `{field_name}` must contain a `type` string")
        })?;

    match field_type {
        "string" => {
            if let Some(pattern) = field_object.get("pattern").and_then(Value::as_str) {
                let passed = matches_pattern(raw_value, pattern)?;
                log_validation_result(
                    verbosity,
                    row_number,
                    record_label,
                    &format!("{field_name} matches pattern"),
                    if passed {
                        ValidationOutcome::Pass
                    } else {
                        ValidationOutcome::Fail
                    },
                );
                if !passed {
                    record_errors.push(format!(
                        "Row {row_number}: `{field_name}` value `{raw_value}` does not match pattern `{pattern}`"
                    ));
                }
            } else {
                log_validation_result(
                    verbosity,
                    row_number,
                    record_label,
                    &format!("{field_name} string"),
                    ValidationOutcome::Pass,
                );
            }
        }
        "number" => {
            let parsed = match raw_value.parse::<f64>() {
                Ok(value) => value,
                Err(_) => {
                    log_validation_result(
                        verbosity,
                        row_number,
                        record_label,
                        &format!("{field_name} number"),
                        ValidationOutcome::Fail,
                    );
                    record_errors.push(format!(
                        "Row {row_number}: `{field_name}` value `{raw_value}` must be a number"
                    ));
                    return Ok(());
                }
            };
            log_validation_result(
                verbosity,
                row_number,
                record_label,
                &format!("{field_name} number"),
                ValidationOutcome::Pass,
            );

            if let Some(minimum) = field_object.get("min").and_then(Value::as_f64) {
                let passed = parsed >= minimum;
                log_validation_result(
                    verbosity,
                    row_number,
                    record_label,
                    &format!("{field_name} >= min"),
                    if passed {
                        ValidationOutcome::Pass
                    } else {
                        ValidationOutcome::Fail
                    },
                );
                if !passed {
                    record_errors.push(format!(
                        "Row {row_number}: `{field_name}` value `{parsed}` is below minimum `{minimum}`"
                    ));
                }
            }

            if let Some(maximum) = field_object.get("max").and_then(Value::as_f64) {
                let passed = parsed <= maximum;
                log_validation_result(
                    verbosity,
                    row_number,
                    record_label,
                    &format!("{field_name} <= max"),
                    if passed {
                        ValidationOutcome::Pass
                    } else {
                        ValidationOutcome::Fail
                    },
                );
                if !passed {
                    record_errors.push(format!(
                        "Row {row_number}: `{field_name}` value `{parsed}` is above maximum `{maximum}`"
                    ));
                }
            }
        }
        "boolean" => {
            let passed = parse_json_bool(raw_value).is_some();
            log_validation_result(
                verbosity,
                row_number,
                record_label,
                &format!("{field_name} boolean"),
                if passed {
                    ValidationOutcome::Pass
                } else {
                    ValidationOutcome::Fail
                },
            );
            if !passed {
                record_errors.push(format!(
                    "Row {row_number}: `{field_name}` value `{raw_value}` must be a boolean"
                ));
            }
        }
        other => {
            return Err(format!(
                "Unsupported field type `{other}` for field `{field_name}`"
            ));
        }
    }

    Ok(())
}

fn validate_panel_engineering(
    row: &Map<String, Value>,
    row_number: usize,
    record_label: &str,
    verbosity: u8,
    record_errors: &mut Vec<String>,
) -> Result<(), String> {
    if let Some(stc_w) = parse_row_number(row, "STC W") {
        let passed = stc_w > 0.0;
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            "STC W > 0",
            if passed {
                ValidationOutcome::Pass
            } else {
                ValidationOutcome::Fail
            },
        );
        if !passed {
            record_errors.push(format!("Row {row_number}: `STC W` must be greater than 0"));
        }
    }

    if let (Some(voc_v), Some(vmp_v)) = (
        parse_row_number(row, "Voc V"),
        parse_row_number(row, "Vmp V"),
    ) {
        let passed = voc_v > vmp_v;
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            "Voc V > Vmp V",
            if passed {
                ValidationOutcome::Pass
            } else {
                ValidationOutcome::Fail
            },
        );
        if !passed {
            record_errors.push(format!(
                "Row {row_number}: `Voc V` must be greater than `Vmp V`"
            ));
        }
    }

    if let (Some(isc_a), Some(imp_a)) = (
        parse_row_number(row, "Isc A"),
        parse_row_number(row, "Imp A"),
    ) {
        let passed = isc_a >= imp_a;
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            "Isc A >= Imp A",
            if passed {
                ValidationOutcome::Pass
            } else {
                ValidationOutcome::Fail
            },
        );
        if !passed {
            record_errors.push(format!(
                "Row {row_number}: `Isc A` must be greater than or equal to `Imp A`"
            ));
        }
    }

    for field_name in ["Voc Temp %/C", "Pmax Temp %/C"] {
        if let Some(temp_coefficient) = parse_row_number(row, field_name) {
            let passed = temp_coefficient <= 0.0;
            log_validation_result(
                verbosity,
                row_number,
                record_label,
                &format!("{field_name} <= 0"),
                if passed {
                    ValidationOutcome::Pass
                } else {
                    ValidationOutcome::Fail
                },
            );
            if !passed {
                record_errors.push(format!(
                    "Row {row_number}: `{field_name}` should normally be less than or equal to 0"
                ));
            }
        }
    }

    Ok(())
}

fn validate_duplicate_primary_key(
    row: &Map<String, Value>,
    row_number: usize,
    record_label: &str,
    verbosity: u8,
    primary_key: &str,
    record_errors: &mut Vec<String>,
    seen_values: &mut BTreeMap<String, usize>,
) {
    let Some(value) = row.get(primary_key).and_then(Value::as_str).map(str::trim) else {
        return;
    };
    if value.is_empty() {
        return;
    }

    if let Some(previous_row) = seen_values.insert(value.to_string(), row_number) {
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            &format!("duplicate {primary_key}"),
            ValidationOutcome::Fail,
        );
        record_errors.push(format!(
            "Row {row_number}: duplicate `{primary_key}` `{value}` also appears in row {previous_row}"
        ));
    } else {
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            &format!("duplicate {primary_key}"),
            ValidationOutcome::Pass,
        );
    }
}

fn validate_duplicate_panel_pair(
    row: &Map<String, Value>,
    row_number: usize,
    record_label: &str,
    verbosity: u8,
    warnings: &mut Vec<Value>,
    seen_panel_pairs: &mut BTreeMap<(String, String), usize>,
) {
    let Some(manufacturer) = row
        .get("Manufacturer")
        .and_then(Value::as_str)
        .map(str::trim)
    else {
        return;
    };
    let Some(model) = row.get("Model").and_then(Value::as_str).map(str::trim) else {
        return;
    };

    if manufacturer.is_empty() || model.is_empty() {
        return;
    }

    let key = (manufacturer.to_string(), model.to_string());
    if let Some(previous_row) = seen_panel_pairs.insert(key.clone(), row_number) {
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            "duplicate Manufacturer + Model",
            ValidationOutcome::Warn,
        );
        warnings.push(Value::Object(Map::from_iter([
            ("recordNumber".to_string(), Value::from(row_number as u64)),
            ("severity".to_string(), Value::String("warning".to_string())),
            (
                "message".to_string(),
                Value::String(format!(
                    "Duplicate `Manufacturer` + `Model` combination `{}` / `{}` also appears in row {previous_row}",
                    key.0, key.1
                )),
            ),
        ])));
    } else {
        log_validation_result(
            verbosity,
            row_number,
            record_label,
            "duplicate Manufacturer + Model",
            ValidationOutcome::Pass,
        );
    }
}

fn parse_row_number(row: &Map<String, Value>, field_name: &str) -> Option<f64> {
    let Some(raw_value) = row.get(field_name).and_then(Value::as_str).map(str::trim) else {
        return None;
    };

    if raw_value.is_empty() {
        return None;
    }

    raw_value.parse::<f64>().ok()
}

fn validation_record_label(row: &Map<String, Value>, row_number: usize) -> String {
    row.get("ID")
        .and_then(Value::as_str)
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(|value| format!("ID {value}"))
        .unwrap_or_else(|| format!("Row {row_number}"))
}

fn log_validation_result(
    verbosity: u8,
    row_number: usize,
    record_label: &str,
    criterion: &str,
    outcome: ValidationOutcome,
) {
    if verbosity >= 3 {
        let status = match outcome {
            ValidationOutcome::Pass => "PASS",
            ValidationOutcome::Fail => "FAIL",
            ValidationOutcome::Warn => "WARN",
        };
        println!("Row {row_number} [{record_label}] {criterion}: {status}");
    }
}

fn normalize_database_name(value: &str) -> String {
    value.trim().to_ascii_lowercase()
}

fn matches_pattern(value: &str, pattern: &str) -> Result<bool, String> {
    let regex = Regex::new(pattern).map_err(|err| {
        format!("Invalid regex pattern `{pattern}` in schema: {err}")
    })?;
    Ok(regex.is_match(value))
}
