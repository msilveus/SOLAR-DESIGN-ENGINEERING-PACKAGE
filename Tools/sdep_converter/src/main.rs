use std::collections::BTreeMap;
use std::env;
use std::fs;
use std::io::{self, Write};
use std::path::Path;
use std::path::PathBuf;
use std::process;

use csv::StringRecord;
use serde_json::{Map, Value};

mod validation;

pub(crate) type SchemaData = Map<String, Value>;
pub(crate) type InputData = Vec<Map<String, Value>>;
pub(crate) type ValidationReport = BTreeMap<String, Value>;

#[derive(Debug)]
struct ManifestData {
    imports: Vec<PathBuf>,
}

#[derive(Debug, Default, PartialEq, Eq)]
struct CliArgs {
    schema_name: Option<String>,
    schema: Option<SchemaData>,
    input_name: Option<String>,
    input: Option<PathBuf>,
    input_data: Option<InputData>,
    output: Option<PathBuf>,
    verbosity: u8,
    all: bool,
    help: bool,
}

fn main() {
    if let Err(message) = run() {
        eprintln!("{message}");
        eprintln!();
        print_usage();
        process::exit(2);
    }
}

fn run() -> Result<(), String> {
    let args = env::args().skip(1).collect::<Vec<_>>();
    run_with_args(args)
}

fn run_with_args(args: Vec<String>) -> Result<(), String> {
    if args.is_empty() {
        print_usage();
        return Ok(());
    }

    match parse_args(args.into_iter()) {
        Ok(args) => {
            if args.help {
                print_usage();
                return Ok(());
            }

            if args.all {
                process_import_manifest(args.verbosity)?;
            } else if let (Some(schema), Some(input_data)) =
                (args.schema.as_ref(), args.input_data.as_ref())
            {
                if args.verbosity >= 1 {
                    println!("Validating CSV input");
                }
                let validation = validate_csv_data(schema, input_data, args.verbosity)?;
                let passed = validation
                    .get("passed")
                    .and_then(Value::as_bool)
                    .unwrap_or(false);

                if passed {
                    let output = args.output.as_ref().ok_or_else(|| {
                        "Validation passed but no output file was provided".to_string()
                    })?;
                    if !ensure_output_directory(output)? {
                        return Ok(());
                    }
                    write_json_output(output, schema, input_data, args.verbosity)?;
                    if args.verbosity >= 1 {
                        println!("Wrote JSON output to {}", output_file_name(output));
                    }
                } else {
                    println!("{validation:#?}");
                }
            }
            Ok(())
        }
        Err(message) => Err(message),
    }
}

fn parse_args<I>(mut args: I) -> Result<CliArgs, String>
where
    I: Iterator<Item = String>,
{
    let mut parsed = CliArgs::default();
    let mut schema_path: Option<PathBuf> = None;
    let mut input_path: Option<PathBuf> = None;

    while let Some(arg) = args.next() {
        match arg.as_str() {
            "-v1" => {
                parsed.verbosity = parsed.verbosity.max(1);
            }
            "-v2" => {
                parsed.verbosity = parsed.verbosity.max(2);
            }
            "-v3" => {
                parsed.verbosity = parsed.verbosity.max(3);
            }
            "-v4" => {
                parsed.verbosity = parsed.verbosity.max(4);
            }
            "--schema" => {
                schema_path = Some(read_path_value(&mut args, "--schema")?);
            }
            "--input" => {
                input_path = Some(read_path_value(&mut args, "--input")?);
            }
            "--output" => {
                parsed.output = Some(read_path_value(&mut args, "--output")?);
            }
            "--all" => {
                parsed.all = true;
            }
            "--help" | "-h" => {
                parsed.help = true;
            }
            other => {
                return Err(format!("Unknown argument: {other}"));
            }
        }
    }

    if parsed.all && (schema_path.is_some() || input_path.is_some() || parsed.output.is_some()) {
        return Err("--all cannot be combined with --schema, --input, or --output".to_string());
    }

    if let Some(path) = schema_path {
        parsed.schema_name = Some(schema_variable_name(&path)?);
        parsed.schema = Some(read_schema_value(&path, parsed.verbosity)?);
    }

    if let Some(path) = input_path {
        parsed.input_name = Some(schema_variable_name(&path)?);
        parsed.input_data = Some(read_input_value(&path, parsed.verbosity)?);
        parsed.input = Some(path);
    }

    Ok(parsed)
}

fn read_path_value<I>(args: &mut I, flag: &str) -> Result<PathBuf, String>
where
    I: Iterator<Item = String>,
{
    let value = args
        .next()
        .ok_or_else(|| format!("Missing value after {flag}"))?;

    if value.starts_with("--") {
        return Err(format!("Missing value after {flag}"));
    }

    Ok(PathBuf::from(value))
}

fn read_schema_value(path: &Path, verbosity: u8) -> Result<SchemaData, String> {
    let path = resolve_schema_path(path);
    if verbosity >= 1 {
        println!("Reading schema file {}", path.display());
    }
    let contents = fs::read_to_string(&path)
        .map_err(|err| format!("Failed to read schema file {}: {err}", path.display()))?;
    let value: Value = serde_json::from_str(&contents)
        .map_err(|err| format!("Failed to parse schema file {}: {err}", path.display()))?;

    let schema = value.as_object().cloned().ok_or_else(|| {
        format!(
            "Schema file {} must contain a top-level JSON object",
            path.display()
        )
    })?;

    if verbosity >= 4 {
        println!("Loaded schema dictionary: {schema:#?}");
    }

    Ok(schema)
}

fn resolve_schema_path(path: &Path) -> PathBuf {
    if path.exists() {
        return path.to_path_buf();
    }

    let candidate = path
        .to_string_lossy()
        .replace("./schemas/", "./src/schemas/")
        .replace("Schemas/", "src/schemas/");
    let candidate_path = PathBuf::from(candidate);
    if candidate_path.exists() {
        return candidate_path;
    }

    if let Some(exe_dir) = executable_directory() {
        let exe_candidate = exe_dir.join(path);
        if exe_candidate.exists() {
            return exe_candidate;
        }

        let exe_transformed = exe_dir.join(&candidate_path);
        if exe_transformed.exists() {
            return exe_transformed;
        }
    }

    path.to_path_buf()
}

fn read_input_value(path: &Path, verbosity: u8) -> Result<InputData, String> {
    let path = resolve_data_path(path);
    if verbosity >= 1 {
        println!("Reading CSV file {}", path.display());
    }
    let mut reader = csv::Reader::from_path(&path)
        .map_err(|err| format!("Failed to read input file {}: {err}", path.display()))?;
    let headers = reader
        .headers()
        .map_err(|err| format!("Failed to read CSV header from {}: {err}", path.display()))?
        .clone();
    let mut rows = Vec::new();

    for (index, record) in reader.records().enumerate() {
        let record = record.map_err(|err| {
            format!(
                "Failed to read CSV row {} from {}: {err}",
                index + 2,
                path.display()
            )
        })?;
        rows.push(csv_record_to_object(&headers, &record, index + 2, &path, verbosity)?);
    }

    if verbosity >= 2 {
        println!("Read {} CSV records from {}", rows.len(), path.display());
    }

    Ok(rows)
}

fn resolve_data_path(path: &Path) -> PathBuf {
    if path.exists() {
        return path.to_path_buf();
    }

    if let Some(exe_dir) = executable_directory() {
        let exe_candidate = exe_dir.join(path);
        if exe_candidate.exists() {
            return exe_candidate;
        }
    }

    path.to_path_buf()
}

fn process_import_manifest(verbosity: u8) -> Result<(), String> {
    let manifest_path = locate_manifest_path()?;
    if verbosity >= 1 {
        println!("Reading import manifest {}", manifest_path.display());
    }

    let manifest = read_manifest_value(&manifest_path)?;
    for import_path in manifest.imports {
        let schema_path = resolve_manifest_schema_path(&manifest_path, &import_path);
        if verbosity >= 1 {
            println!("Processing schema {}", schema_path.display());
        }

        let schema = read_schema_value(&schema_path, verbosity)?;
        let input_path = schema
            .get("input")
            .and_then(Value::as_str)
            .map(PathBuf::from)
            .ok_or_else(|| {
                format!(
                    "Schema file {} must contain an `input` string",
                    schema_path.display()
                )
            })?;
        let output_path = schema
            .get("output")
            .and_then(Value::as_str)
            .map(PathBuf::from)
            .ok_or_else(|| {
                format!(
                    "Schema file {} must contain an `output` string",
                    schema_path.display()
                )
            })?;

        if verbosity >= 1 {
            println!("Using input {}", input_path.display());
            println!("Using output {}", output_path.display());
        }

        if !ensure_output_directory(&output_path)? {
            continue;
        }

        let input_data = read_input_value(&input_path, verbosity)?;
        let validation = validate_csv_data(&schema, &input_data, verbosity)?;
        let passed = validation
            .get("passed")
            .and_then(Value::as_bool)
            .unwrap_or(false);

        if passed {
            write_json_output(&output_path, &schema, &input_data, verbosity)?;
            if verbosity >= 1 {
                println!("Wrote JSON output to {}", output_file_name(&output_path));
            }
        } else {
            println!("{validation:#?}");
        }
    }

    Ok(())
}

fn locate_manifest_path() -> Result<PathBuf, String> {
    let candidates = [
        Path::new("./schemas/import_manifest.json"),
        Path::new("./src/schemas/import_manifest.json"),
    ];

    let mut resolved_candidates = candidates
        .into_iter()
        .map(Path::to_path_buf)
        .collect::<Vec<_>>();

    if let Some(exe_dir) = executable_directory() {
        resolved_candidates.push(exe_dir.join("schemas/import_manifest.json"));
        resolved_candidates.push(exe_dir.join("src/schemas/import_manifest.json"));
    }

    resolved_candidates
        .into_iter()
        .find(|path| path.exists())
        .ok_or_else(|| {
            "Could not find import_manifest.json in ./schemas, ./src/schemas, or the application directory".to_string()
        })
}

fn read_manifest_value(path: &Path) -> Result<ManifestData, String> {
    let contents = fs::read_to_string(path)
        .map_err(|err| format!("Failed to read manifest file {}: {err}", path.display()))?;
    let value: Value = serde_json::from_str(&contents)
        .map_err(|err| format!("Failed to parse manifest file {}: {err}", path.display()))?;

    let imports = value
        .get("imports")
        .and_then(Value::as_array)
        .ok_or_else(|| {
            format!(
                "Manifest file {} must contain a top-level `imports` array",
                path.display()
            )
        })?
        .iter()
        .filter_map(Value::as_str)
        .map(PathBuf::from)
        .collect::<Vec<_>>();

    Ok(ManifestData { imports })
}

fn resolve_manifest_schema_path(manifest_path: &Path, import_path: &Path) -> PathBuf {
    let manifest_dir = manifest_path.parent().unwrap_or_else(|| Path::new("."));
    let direct = manifest_dir.join(import_path);
    if direct.exists() {
        return direct;
    }

    let import_string = import_path.to_string_lossy();
    if let Some(stripped) = import_string.strip_prefix("Schemas/") {
        let alt = manifest_dir.join(stripped);
        if alt.exists() {
            return alt;
        }
    }
    if let Some(stripped) = import_string.strip_prefix("Schemas\\") {
        let alt = manifest_dir.join(stripped);
        if alt.exists() {
            return alt;
        }
    }

    import_path.to_path_buf()
}

fn csv_record_to_object(
    headers: &StringRecord,
    record: &StringRecord,
    line_number: usize,
    path: &Path,
    verbosity: u8,
) -> Result<Map<String, Value>, String> {
    if record.len() > headers.len() {
        return Err(format!(
            "CSV row {} in {} has more fields than the header row",
            line_number,
            path.display()
        ));
    }

    let mut row = Map::with_capacity(headers.len());

    for (index, header) in headers.iter().enumerate() {
        let value = record.get(index).unwrap_or("");
        row.insert(header.to_string(), Value::String(value.to_string()));
    }

    if verbosity >= 4 {
        println!("Created CSV dictionary for row {line_number}: {row:#?}");
    }

    Ok(row)
}

fn validate_csv_data(
    schema: &SchemaData,
    input_data: &InputData,
    verbosity: u8,
) -> Result<ValidationReport, String> {
    validation::validate_database(schema, input_data, verbosity)
}

fn write_json_output(
    path: &Path,
    schema: &SchemaData,
    input_data: &InputData,
    verbosity: u8,
) -> Result<(), String> {
    if verbosity >= 1 {
        println!("Writing JSON file {}", path.display());
    }
    let json_value = convert_input_data_to_json(schema, input_data, verbosity)?;
    if verbosity >= 2 {
        println!("Converted {} records to JSON", input_data.len());
    }
    let json = serde_json::to_string_pretty(&json_value)
        .map_err(|err| format!("Failed to serialize JSON output {}: {err}", path.display()))?;

    fs::write(path, json)
        .map_err(|err| format!("Failed to write JSON output {}: {err}", path.display()))
}

fn convert_input_data_to_json(
    schema: &SchemaData,
    input_data: &InputData,
    verbosity: u8,
) -> Result<Value, String> {
    let field_definitions = schema
        .get("fields")
        .and_then(Value::as_object)
        .ok_or_else(|| "Schema is missing a top-level `fields` object".to_string())?;

    let database_name = schema
        .get("database")
        .or_else(|| schema.get("name"))
        .and_then(Value::as_str)
        .ok_or_else(|| "Schema is missing a top-level `database` string".to_string())?;
    let version = schema
        .get("version")
        .cloned()
        .ok_or_else(|| "Schema is missing a top-level `version` value".to_string())?;
    let mut rows = Vec::with_capacity(input_data.len());

    for row in input_data {
        let mut json_row = Map::with_capacity(row.len());

        for (field_name, field_value) in row {
            let converted_value = match field_definitions.get(field_name) {
                Some(field_definition) => {
                    convert_field_value(field_name, field_definition, field_value)?
                }
                None => field_value.clone(),
            };
            json_row.insert(field_name.clone(), converted_value);
        }

        if verbosity >= 4 {
            println!("Created JSON dictionary for converted row: {json_row:#?}");
        }
        rows.push(Value::Object(json_row));
    }

    let header = Map::from_iter([
        (
            "databaseName".to_string(),
            Value::String(database_name.to_string()),
        ),
        ("version".to_string(), version),
        (
            "recordCount".to_string(),
            Value::from(input_data.len() as u64),
        ),
    ]);

    if verbosity >= 4 {
        println!("Created JSON header dictionary: {header:#?}");
    }

    Ok(Value::Object(Map::from_iter([
        (
            "header".to_string(),
            Value::Object(header),
        ),
        ("records".to_string(), Value::Array(rows)),
    ])))
}

fn convert_field_value(
    field_name: &str,
    field_definition: &Value,
    field_value: &Value,
) -> Result<Value, String> {
    let field_object = field_definition
        .as_object()
        .ok_or_else(|| format!("Field definition for `{field_name}` must be a JSON object"))?;
    let raw_value = field_value.as_str().unwrap_or("").trim();

    if raw_value.is_empty() {
        return Ok(Value::String(String::new()));
    }

    let field_type = field_object
        .get("type")
        .and_then(Value::as_str)
        .ok_or_else(|| {
            format!("Field definition for `{field_name}` must contain a `type` string")
        })?;

    match field_type {
        "string" => Ok(Value::String(raw_value.to_string())),
        "number" => parse_json_number(raw_value)
            .map(Value::Number)
            .ok_or_else(|| {
                format!(
                    "Field `{field_name}` value `{raw_value}` could not be converted to a number"
                )
            }),
        "boolean" => parse_json_bool(raw_value).map(Value::Bool).ok_or_else(|| {
            format!("Field `{field_name}` value `{raw_value}` could not be converted to a boolean")
        }),
        other => Err(format!(
            "Unsupported field type `{other}` for field `{field_name}`"
        )),
    }
}

fn parse_json_number(raw_value: &str) -> Option<serde_json::Number> {
    if let Ok(value) = raw_value.parse::<i64>() {
        return Some(serde_json::Number::from(value));
    }

    raw_value
        .parse::<f64>()
        .ok()
        .and_then(serde_json::Number::from_f64)
}

pub(crate) fn parse_json_bool(raw_value: &str) -> Option<bool> {
    match raw_value.to_ascii_lowercase().as_str() {
        "true" => Some(true),
        "false" => Some(false),
        _ => None,
    }
}

fn ensure_output_directory(path: &Path) -> Result<bool, String> {
    let Some(parent) = path.parent() else {
        return Ok(true);
    };

    if parent.as_os_str().is_empty() || parent.exists() {
        return Ok(true);
    }

    let folder_name = parent
        .file_name()
        .and_then(|value| value.to_str())
        .map(|value| value.to_string())
        .unwrap_or_else(|| parent.display().to_string());

    loop {
        print!(
            "Output folder '{}' does not exist. Type Yes to create it or No to exit: ",
            folder_name
        );
        io::stdout()
            .flush()
            .map_err(|err| format!("Failed to write prompt: {err}"))?;

        let mut response = String::new();
        io::stdin()
            .read_line(&mut response)
            .map_err(|err| format!("Failed to read response: {err}"))?;

        match parse_yes_no_response(&response) {
            Some(true) => {
                fs::create_dir_all(parent).map_err(|err| {
                    format!("Failed to create output folder {}: {err}", parent.display())
                })?;
                return Ok(true);
            }
            Some(false) => return Ok(false),
            _ => {
                println!("Please respond with Yes or No.");
            }
        }
    }
}

fn parse_yes_no_response(response: &str) -> Option<bool> {
    match response.trim().to_ascii_lowercase().as_str() {
        "yes" => Some(true),
        "no" => Some(false),
        _ => None,
    }
}

fn output_file_name(path: &Path) -> String {
    path.file_name()
        .and_then(|value| value.to_str())
        .map(|value| value.to_string())
        .unwrap_or_else(|| path.display().to_string())
}

fn schema_variable_name(path: &Path) -> Result<String, String> {
    let stem = path
        .file_stem()
        .and_then(|value| value.to_str())
        .ok_or_else(|| format!("Schema file {} must have a valid file name", path.display()))?;

    Ok(stem.to_string())
}

fn executable_directory() -> Option<PathBuf> {
    env::current_exe()
        .ok()
        .and_then(|path| path.parent().map(Path::to_path_buf))
}

fn print_usage() {
    eprintln!("Usage:");
    eprintln!("  sdep_converter --schema <schema.json> --input <input.csv> --output <output.json>");
    eprintln!("  sdep_converter --all");
    eprintln!();
    eprintln!("Examples:");
    eprintln!(
        "  Single data conversion:\n    sdep_converter --schema ./schemas/panels_schema.json --input ./Data/panels.csv --output ./Data/panels.json"
    );
    eprintln!(
        "  Read the manifest and process all data conversions:\n    sdep_converter --all"
    );
    eprintln!();
    eprintln!("Options:");
    eprintln!("  -h, --help       Show this help text");
    eprintln!("  --schema <file>  Load a single schema JSON file");
    eprintln!("  --input <file>   Load a single CSV input file");
    eprintln!("  --output <file>  Write the converted JSON output file");
    eprintln!("  --all            Read the import manifest and process every conversion");
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::env;
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn parse(items: &[&str]) -> Result<CliArgs, String> {
        parse_args(items.iter().map(|s| s.to_string()))
    }

    fn write_temp_schema(contents: &str) -> PathBuf {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock should be monotonic")
            .as_nanos();
        let path = env::temp_dir().join(format!("sdep_converter_schema_{stamp}.json"));
        fs::write(&path, contents).expect("should write temp schema");
        path
    }

    fn write_temp_csv(contents: &str) -> PathBuf {
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock should be monotonic")
            .as_nanos();
        let path = env::temp_dir().join(format!("sdep_converter_input_{stamp}.csv"));
        fs::write(&path, contents).expect("should write temp csv");
        path
    }

    #[test]
    fn parses_schema_input_output() {
        let schema_path = write_temp_schema(r#"{"name":"Panels","primaryKey":"ID"}"#);
        let schema_path_str = schema_path.to_string_lossy().to_string();
        let input_path = write_temp_csv("ID,Manufacturer\nP000001,ACME\nP000002,SolarCo\n");
        let input_path_str = input_path.to_string_lossy().to_string();
        let args = parse(&[
            "--schema",
            &schema_path_str,
            "--input",
            &input_path_str,
            "--output",
            "out.json",
        ])
        .expect("parse should succeed");
        let schema_name = schema_path
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap();
        let input_name = input_path
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap();

        assert_eq!(
            args,
            CliArgs {
                schema_name: Some(schema_name.to_string()),
                schema: Some(serde_json::Map::from_iter([
                    ("name".to_string(), Value::String("Panels".to_string())),
                    ("primaryKey".to_string(), Value::String("ID".to_string())),
                ])),
                input_name: Some(input_name.to_string()),
                input: Some(PathBuf::from(&input_path_str)),
                input_data: Some(vec![
                    serde_json::Map::from_iter([
                        ("ID".to_string(), Value::String("P000001".to_string())),
                        (
                            "Manufacturer".to_string(),
                            Value::String("ACME".to_string())
                        ),
                    ]),
                    serde_json::Map::from_iter([
                        ("ID".to_string(), Value::String("P000002".to_string())),
                        (
                            "Manufacturer".to_string(),
                            Value::String("SolarCo".to_string())
                        ),
                    ]),
                ]),
                output: Some(PathBuf::from("out.json")),
                verbosity: 0,
                all: false,
                help: false,
            }
        );
    }

    #[test]
    fn parses_input_csv_file() {
        let input_path = write_temp_csv("ID,Manufacturer\nP000001,ACME\nP000002,SolarCo\n");
        let input_path_str = input_path.to_string_lossy().to_string();
        let args = parse(&["--input", &input_path_str]).expect("parse should succeed");
        let input_name = input_path
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap();

        assert_eq!(args.input_name.as_deref(), Some(input_name));
        assert_eq!(
            args.input_data,
            Some(vec![
                serde_json::Map::from_iter([
                    ("ID".to_string(), Value::String("P000001".to_string())),
                    (
                        "Manufacturer".to_string(),
                        Value::String("ACME".to_string())
                    ),
                ]),
                serde_json::Map::from_iter([
                    ("ID".to_string(), Value::String("P000002".to_string())),
                    (
                        "Manufacturer".to_string(),
                        Value::String("SolarCo".to_string())
                    ),
                ]),
            ])
        );
    }

    #[test]
    fn parses_verbosity_flags() {
        let args = parse(&["-v1", "-v3", "-v2"]).expect("parse should succeed");
        assert_eq!(args.verbosity, 3);
    }

    #[test]
    fn reads_import_manifest_entries() {
        let manifest_path = locate_manifest_path().expect("manifest path should exist");
        let manifest = read_manifest_value(&manifest_path)
            .expect("manifest should parse");
        assert_eq!(manifest.imports.len(), 3);
    }

    #[test]
    fn resolves_manifest_schema_path_from_schemas_prefix() {
        let manifest_path = locate_manifest_path().expect("manifest path should exist");
        let resolved =
            resolve_manifest_schema_path(&manifest_path, Path::new("Schemas/panels_schema.json"));
        let expected = manifest_path
            .parent()
            .unwrap_or_else(|| Path::new("."))
            .join("panels_schema.json");
        assert_eq!(resolved, expected);
    }

    #[test]
    fn validates_csv_data_against_schema_fields() {
        let schema = serde_json::Map::from_iter([
            ("name".to_string(), Value::String("Panels".to_string())),
            (
                "requiredFields".to_string(),
                Value::Array(vec![
                    Value::String("ID".to_string()),
                    Value::String("Manufacturer".to_string()),
                    Value::String("Wattage".to_string()),
                ]),
            ),
            (
                "fields".to_string(),
                Value::Object(serde_json::Map::from_iter([
                    (
                        "ID".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("string".to_string())),
                            (
                                "pattern".to_string(),
                                Value::String("^P[0-9]{6}$".to_string()),
                            ),
                        ])),
                    ),
                    (
                        "Manufacturer".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("string".to_string())),
                            ("required".to_string(), Value::Bool(true)),
                        ])),
                    ),
                    (
                        "Wattage".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(1)),
                        ])),
                    ),
                ])),
            ),
        ]);
        let input_data = vec![serde_json::Map::from_iter([
            ("ID".to_string(), Value::String("BAD".to_string())),
            ("Manufacturer".to_string(), Value::String(String::new())),
            ("Wattage".to_string(), Value::String("0".to_string())),
        ])];

        let validation =
            validate_csv_data(&schema, &input_data, 0).expect("validation should succeed");

        let record = validation
            .get("record_1")
            .and_then(Value::as_object)
            .expect("record_1 should be present");
        let errors = record
            .get("errors")
            .and_then(Value::as_array)
            .expect("errors should be present");

        assert_eq!(validation.get("passed"), Some(&Value::Bool(false)));
        assert!(!validation.contains_key("validatedRecords"));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("does not match pattern"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("is required"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("below minimum"))
                .unwrap_or(false)
        }));
    }

    #[test]
    fn validates_csv_data_reports_success_count() {
        let schema = serde_json::Map::from_iter([
            ("name".to_string(), Value::String("Panels".to_string())),
            (
                "requiredFields".to_string(),
                Value::Array(vec![
                    Value::String("ID".to_string()),
                    Value::String("Manufacturer".to_string()),
                ]),
            ),
            (
                "fields".to_string(),
                Value::Object(serde_json::Map::from_iter([
                    (
                        "ID".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("string".to_string())),
                            (
                                "pattern".to_string(),
                                Value::String("^P[0-9]{6}$".to_string()),
                            ),
                        ])),
                    ),
                    (
                        "Manufacturer".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("string".to_string())),
                            ("required".to_string(), Value::Bool(true)),
                        ])),
                    ),
                ])),
            ),
        ]);
        let input_data = vec![
            serde_json::Map::from_iter([
                ("ID".to_string(), Value::String("P000001".to_string())),
                (
                    "Manufacturer".to_string(),
                    Value::String("ACME".to_string()),
                ),
            ]),
            serde_json::Map::from_iter([
                ("ID".to_string(), Value::String("P000002".to_string())),
                (
                    "Manufacturer".to_string(),
                    Value::String("SolarCo".to_string()),
                ),
            ]),
        ];

        let validation =
            validate_csv_data(&schema, &input_data, 0).expect("validation should succeed");

        assert_eq!(validation.get("passed"), Some(&Value::Bool(true)));
        assert_eq!(
            validation.get("validatedRecords"),
            Some(&Value::from(2_u64))
        );
        assert_eq!(validation.len(), 2);
    }

    #[test]
    fn validates_panel_engineering_rules_and_schema_ranges() {
        let schema = serde_json::Map::from_iter([
            ("database".to_string(), Value::String("Panels".to_string())),
            ("version".to_string(), Value::from(1.0)),
            (
                "requiredFields".to_string(),
                Value::Array(vec![
                    Value::String("ID".to_string()),
                    Value::String("Manufacturer".to_string()),
                    Value::String("Model".to_string()),
                    Value::String("STC W".to_string()),
                    Value::String("Voc V".to_string()),
                    Value::String("Vmp V".to_string()),
                    Value::String("Isc A".to_string()),
                    Value::String("Imp A".to_string()),
                    Value::String("Voc Temp %/C".to_string()),
                ]),
            ),
            (
                "fields".to_string(),
                Value::Object(serde_json::Map::from_iter([
                    (
                        "ID".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "Manufacturer".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "Model".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "STC W".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(1)),
                        ])),
                    ),
                    (
                        "Voc V".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Vmp V".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Isc A".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Imp A".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Voc Temp %/C".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("max".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Pmax Temp %/C".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("max".to_string(), Value::from(0)),
                        ])),
                    ),
                    (
                        "Length mm".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(1000)),
                            ("max".to_string(), Value::from(2200)),
                        ])),
                    ),
                    (
                        "Weight lb".to_string(),
                        Value::Object(serde_json::Map::from_iter([
                            ("type".to_string(), Value::String("number".to_string())),
                            ("min".to_string(), Value::from(20)),
                            ("max".to_string(), Value::from(80)),
                        ])),
                    ),
                ])),
            ),
        ]);
        let input_data = vec![serde_json::Map::from_iter([
            ("ID".to_string(), Value::String("P000001".to_string())),
            (
                "Manufacturer".to_string(),
                Value::String("ACME".to_string()),
            ),
            ("Model".to_string(), Value::String("Alpha".to_string())),
            ("STC W".to_string(), Value::String("0".to_string())),
            ("Voc V".to_string(), Value::String("30".to_string())),
            ("Vmp V".to_string(), Value::String("35".to_string())),
            ("Isc A".to_string(), Value::String("8".to_string())),
            ("Imp A".to_string(), Value::String("9".to_string())),
            ("Voc Temp %/C".to_string(), Value::String("0.1".to_string())),
            ("Pmax Temp %/C".to_string(), Value::String("0.2".to_string())),
            ("Length mm".to_string(), Value::String("900".to_string())),
            ("Weight lb".to_string(), Value::String("90".to_string())),
        ])];

        let validation =
            validate_csv_data(&schema, &input_data, 0).expect("validation should succeed");
        let record = validation
            .get("record_1")
            .and_then(Value::as_object)
            .expect("record_1 should be present");
        let errors = record
            .get("errors")
            .and_then(Value::as_array)
            .expect("errors should be present");

        assert_eq!(validation.get("passed"), Some(&Value::Bool(false)));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("greater than 0"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("greater than `Vmp V`"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("greater than or equal to `Imp A`"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("less than or equal to 0"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("below minimum"))
                .unwrap_or(false)
        }));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("above maximum"))
                .unwrap_or(false)
        }));
    }

    #[test]
    fn validates_duplicate_ids_and_warns_on_duplicate_panel_pairs() {
        let schema = serde_json::Map::from_iter([
            ("database".to_string(), Value::String("Panels".to_string())),
            ("version".to_string(), Value::from(1.0)),
            (
                "fields".to_string(),
                Value::Object(serde_json::Map::from_iter([
                    (
                        "ID".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "Manufacturer".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "Model".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                ])),
            ),
        ]);
        let input_data = vec![
            serde_json::Map::from_iter([
                ("ID".to_string(), Value::String("P000001".to_string())),
                (
                    "Manufacturer".to_string(),
                    Value::String("ACME".to_string()),
                ),
                ("Model".to_string(), Value::String("Alpha".to_string())),
            ]),
            serde_json::Map::from_iter([
                ("ID".to_string(), Value::String("P000001".to_string())),
                (
                    "Manufacturer".to_string(),
                    Value::String("ACME".to_string()),
                ),
                ("Model".to_string(), Value::String("Alpha".to_string())),
            ]),
        ];

        let validation =
            validate_csv_data(&schema, &input_data, 0).expect("validation should succeed");
        let record = validation
            .get("record_2")
            .and_then(Value::as_object)
            .expect("record_2 should be present");
        let errors = record
            .get("errors")
            .and_then(Value::as_array)
            .expect("errors should be present");
        let warnings = validation
            .get("warnings")
            .and_then(Value::as_array)
            .expect("warnings should be present");

        assert_eq!(validation.get("passed"), Some(&Value::Bool(false)));
        assert!(errors.iter().any(|message| {
            message
                .as_str()
                .map(|text| text.contains("duplicate `ID`"))
                .unwrap_or(false)
        }));
        assert!(warnings.iter().any(|warning| {
            warning
                .as_object()
                .and_then(|object| object.get("message"))
                .and_then(Value::as_str)
                .map(|text| text.contains("Duplicate `Manufacturer` + `Model`"))
                .unwrap_or(false)
        }));
    }

    #[test]
    fn writes_json_output_file() {
        let output_path = env::temp_dir().join(format!(
            "sdep_converter_output_{}.json",
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .expect("clock should be monotonic")
                .as_nanos()
        ));
        let schema = serde_json::Map::from_iter([
            (
                "database".to_string(),
                Value::String("Panels".to_string()),
            ),
            ("version".to_string(), Value::from(1.0)),
            (
                "fields".to_string(),
                Value::Object(serde_json::Map::from_iter([
                    (
                        "ID".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("string".to_string()),
                        )])),
                    ),
                    (
                        "Wattage".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("number".to_string()),
                        )])),
                    ),
                    (
                        "Installed".to_string(),
                        Value::Object(serde_json::Map::from_iter([(
                            "type".to_string(),
                            Value::String("boolean".to_string()),
                        )])),
                    ),
                ])),
            ),
        ]);
        let input_data = vec![serde_json::Map::from_iter([
            ("ID".to_string(), Value::String("P000001".to_string())),
            ("Wattage".to_string(), Value::String("250".to_string())),
            ("Installed".to_string(), Value::String("true".to_string())),
        ])];

        write_json_output(&output_path, &schema, &input_data, 0).expect("write should succeed");

        let written = fs::read_to_string(&output_path).expect("file should exist");
        assert!(written.contains("\"header\""));
        assert!(written.contains("\"databaseName\": \"Panels\""));
        assert!(written.contains("\"version\": 1.0"));
        assert!(written.contains("\"recordCount\": 1"));
        assert!(written.contains("\"records\""));
        assert!(written.contains("\"ID\": \"P000001\""));
        assert!(written.contains("\"Wattage\": 250"));
        assert!(written.contains("\"Installed\": true"));
    }

    #[test]
    fn parses_yes_no_responses_case_insensitively() {
        assert_eq!(parse_yes_no_response("YES"), Some(true));
        assert_eq!(parse_yes_no_response(" yes "), Some(true));
        assert_eq!(parse_yes_no_response("No"), Some(false));
        assert_eq!(parse_yes_no_response("maybe"), None);
    }

    #[test]
    fn parses_all_flag() {
        let args = parse(&["--all"]).expect("parse should succeed");
        assert!(args.all);
        assert!(!args.help);
    }

    #[test]
    fn parses_help_flag() {
        let args = parse(&["--help"]).expect("parse should succeed");
        assert!(args.help);
    }

    #[test]
    fn run_with_no_args_succeeds() {
        assert!(run_with_args(Vec::new()).is_ok());
    }

    #[test]
    fn rejects_missing_value() {
        let err = parse(&["--input"]).expect_err("parse should fail");
        assert_eq!(err, "Missing value after --input");
    }

    #[test]
    fn rejects_combined_all_and_input() {
        let err = parse(&["--all", "--input", "in.csv"]).expect_err("parse should fail");
        assert_eq!(err, "--all cannot be combined with --schema, --input, or --output");
    }
}
