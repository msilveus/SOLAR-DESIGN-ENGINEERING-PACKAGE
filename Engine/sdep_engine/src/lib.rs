use serde_json::Value;
use std::collections::HashMap;
use std::ffi::c_char;
use std::fs;
use std::path::{Path, PathBuf};
use std::ptr;
use std::sync::{Arc, OnceLock, RwLock};

#[cfg(windows)]
use std::ffi::OsString;
#[cfg(windows)]
use std::os::windows::ffi::OsStringExt;

#[cfg(windows)]
const GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS: u32 = 0x0000_0004;
#[cfg(windows)]
const GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT: u32 = 0x0000_0002;

#[cfg(windows)]
type Hmodule = *mut core::ffi::c_void;

#[cfg(windows)]
#[link(name = "kernel32")]
unsafe extern "system" {
    fn GetModuleHandleExW(
        dwFlags: u32,
        lpModuleName: *const u16,
        phModule: *mut Hmodule,
    ) -> i32;
    fn GetModuleFileNameW(hModule: Hmodule, lpFilename: *mut u16, nSize: u32) -> u32;
}

pub const SDEP_SUCCESS: i32 = 0;
pub const SDEP_BAD_ID: i32 = 1;
pub const SDEP_BUFFER_TOO_SMALL: i32 = 2;
pub const SDEP_INVALID_ARGUMENT: i32 = 3;
pub const SDEP_DATABASE_NOT_LOADED: i32 = 4;
pub const SDEP_DATABASE_CORRUPT: i32 = 5;
pub const SDEP_WRONG_DATABASE_VERSION: i32 = 6;
pub const SDEP_INTERNAL_ERROR: i32 = 1000;

#[derive(Clone, Debug)]
struct Record {
    fields: HashMap<String, Value>,
}

#[derive(Clone, Debug)]
struct Dataset {
    ids: Vec<String>,
    records_by_id: HashMap<String, Record>,
}

#[derive(Clone, Debug)]
struct DatabaseStore {
    datasets: HashMap<String, Dataset>,
}

#[derive(Clone, Debug)]
enum EngineState {
    Unloaded,
    Loaded(Arc<DatabaseStore>),
    Failed(i32),
}

static STORE: OnceLock<RwLock<EngineState>> = OnceLock::new();

fn store_cell() -> &'static RwLock<EngineState> {
    STORE.get_or_init(|| RwLock::new(EngineState::Unloaded))
}

fn set_store(store: DatabaseStore) {
    if let Ok(mut guard) = store_cell().write() {
        *guard = EngineState::Loaded(Arc::new(store));
    }
}

fn set_failure(code: i32) {
    if let Ok(mut guard) = store_cell().write() {
        *guard = EngineState::Failed(code);
    }
}

fn current_state() -> EngineState {
    store_cell()
        .read()
        .map(|guard| guard.clone())
        .unwrap_or(EngineState::Failed(SDEP_INTERNAL_ERROR))
}

fn ensure_store() -> Result<Arc<DatabaseStore>, i32> {
    match current_state() {
        EngineState::Loaded(store) => Ok(store),
        EngineState::Failed(code) => Err(code),
        EngineState::Unloaded => {
            let dir = dll_directory()
                .map(|path| path.join("DataJson"))
                .unwrap_or_else(default_data_dir);

            match load_store_from_dir(&dir) {
                Ok(store) => {
                    set_store(store);
                    match current_state() {
                        EngineState::Loaded(store) => Ok(store),
                        EngineState::Failed(code) => Err(code),
                        EngineState::Unloaded => Err(SDEP_INTERNAL_ERROR),
                    }
                }
                Err(code) => {
                    set_failure(code);
                    Err(code)
                }
            }
        }
    }
}

fn default_data_dir() -> PathBuf {
    std::env::current_dir()
        .unwrap_or_else(|_| PathBuf::from("."))
        .join("DataJson")
}

#[cfg(windows)]
fn dll_directory() -> Option<PathBuf> {
    let mut module: Hmodule = ptr::null_mut();
    let address = sdep_init as *const () as *const u16;
    let ok = unsafe {
        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            address,
            &mut module,
        )
    };
    if ok == 0 || module.is_null() {
        return None;
    }

    let mut buffer = vec![0u16; 260];
    let len = unsafe { GetModuleFileNameW(module, buffer.as_mut_ptr(), buffer.len() as u32) };
    if len == 0 {
        return None;
    }

    buffer.truncate(len as usize);
    let module_path = OsString::from_wide(&buffer);
    let mut path = PathBuf::from(module_path);
    path.pop();
    Some(path)
}

#[cfg(not(windows))]
fn dll_directory() -> Option<PathBuf> {
    None
}

fn load_store_from_dir(dir: &Path) -> Result<DatabaseStore, i32> {
    let dir = resolve_data_dir(dir);
    if !dir.is_dir() {
        return Err(SDEP_DATABASE_NOT_LOADED);
    }

    let mut datasets = HashMap::new();
    let entries = fs::read_dir(&dir).map_err(|_| SDEP_INTERNAL_ERROR)?;
    let mut file_count = 0usize;

    for entry in entries {
        let entry = entry.map_err(|_| SDEP_INTERNAL_ERROR)?;
        let path = entry.path();
        if !path.is_file() {
            continue;
        }

        let is_json = path
            .extension()
            .and_then(|ext| ext.to_str())
            .map(|ext| ext.eq_ignore_ascii_case("json"))
            .unwrap_or(false);
        if !is_json {
            continue;
        }

        file_count += 1;

        let text = fs::read_to_string(&path).map_err(|_| SDEP_INTERNAL_ERROR)?;
        let json: Value = serde_json::from_str(&text).map_err(|_| SDEP_DATABASE_CORRUPT)?;

        let dataset = parse_dataset(&json, &path)?;
        datasets.insert(dataset.0, dataset.1);
    }

    if file_count == 0 {
        return Err(SDEP_DATABASE_NOT_LOADED);
    }

    Ok(DatabaseStore { datasets })
}

fn resolve_data_dir(path: &Path) -> PathBuf {
    if path
        .file_name()
        .and_then(|name| name.to_str())
        .map(|name| name.eq_ignore_ascii_case("DataJson"))
        .unwrap_or(false)
    {
        return path.to_path_buf();
    }

    path.join("DataJson")
}

fn parse_dataset(json: &Value, _path: &Path) -> Result<(String, Dataset), i32> {
    let header = json
        .get("header")
        .and_then(Value::as_object)
        .ok_or(SDEP_DATABASE_CORRUPT)?;
    let database_name = header
        .get("databaseName")
        .and_then(Value::as_str)
        .map(normalize_name)
        .filter(|name| !name.is_empty())
        .or_else(|| {
            _path
                .file_stem()
                .and_then(|stem| stem.to_str())
                .map(normalize_name)
                .filter(|name| !name.is_empty())
        })
        .ok_or(SDEP_DATABASE_CORRUPT)?;

    let version = header
        .get("version")
        .and_then(Value::as_f64)
        .ok_or(SDEP_DATABASE_CORRUPT)?;
    if (version - 1.0).abs() > f64::EPSILON {
        return Err(SDEP_WRONG_DATABASE_VERSION);
    }

    let declared_count = header
        .get("recordCount")
        .and_then(Value::as_u64)
        .ok_or(SDEP_DATABASE_CORRUPT)? as usize;

    let records = json
        .get("records")
        .and_then(Value::as_array)
        .ok_or(SDEP_DATABASE_CORRUPT)?;
    if records.len() != declared_count {
        return Err(SDEP_DATABASE_CORRUPT);
    }

    let mut ids = Vec::with_capacity(records.len());
    let mut records_by_id = HashMap::with_capacity(records.len());

    for record in records {
        let object = record
            .as_object()
            .ok_or(SDEP_DATABASE_CORRUPT)?;

        let mut fields = HashMap::with_capacity(object.len());
        let mut id_value = None;

        for (key, value) in object {
            let normalized_key = normalize_name(key);
            if normalized_key.is_empty() {
                continue;
            }

            if normalized_key == "id" {
                id_value = value.as_str().map(|s| s.to_owned()).or_else(|| value_to_string(value));
            }

            fields.insert(normalized_key, value.clone());
        }

        let id = id_value.ok_or(SDEP_DATABASE_CORRUPT)?;
        if records_by_id.contains_key(&id) {
            return Err(SDEP_DATABASE_CORRUPT);
        }
        ids.push(id.clone());
        records_by_id.insert(id, Record { fields });
    }

    Ok((database_name, Dataset { ids, records_by_id }))
}

fn normalize_name(value: &str) -> String {
    let lowered = value.trim().to_ascii_lowercase();
    if lowered == "id" {
        return "id".to_owned();
    }

    match lowered.as_str() {
        "vmp v" => "vmp".to_owned(),
        "voc v" => "voc".to_owned(),
        "imp a" => "imp".to_owned(),
        "isc a" => "isc".to_owned(),
        _ => {
            let mut out = String::with_capacity(lowered.len());
            let mut previous_underscore = false;
            for ch in lowered.chars() {
                let mapped = if ch.is_ascii_alphanumeric() {
                    ch
                } else {
                    '_'
                };
                if mapped == '_' {
                    if !previous_underscore {
                        out.push(mapped);
                        previous_underscore = true;
                    }
                } else {
                    out.push(mapped);
                    previous_underscore = false;
                }
            }
            out.trim_matches('_').to_owned()
        }
    }
}

fn value_to_string(value: &Value) -> Option<String> {
    match value {
        Value::String(text) => Some(text.clone()),
        Value::Number(number) => Some(number.to_string()),
        Value::Bool(flag) => Some(flag.to_string()),
        Value::Null => Some(String::new()),
        _ => None,
    }
}

fn lookup_record<'a>(store: &'a DatabaseStore, database: &str, id: &str) -> Option<&'a Record> {
    store
        .datasets
        .get(database)
        .and_then(|dataset| dataset.records_by_id.get(id))
}

pub(crate) fn lookup_numeric(
    database: &str,
    field: &str,
    id: *const c_char,
    out: *mut f64,
) -> i32 {
    if out.is_null() {
        return SDEP_INVALID_ARGUMENT;
    }

    let store = match ensure_store() {
        Ok(store) => store,
        Err(code) => return code,
    };
    let Some(id) = c_string_arg(id) else {
        return SDEP_BAD_ID;
    };

    let Some(record) = lookup_record(&store, database, &id) else {
        return SDEP_BAD_ID;
    };
    let Some(value) = record.fields.get(field) else {
        return SDEP_INTERNAL_ERROR;
    };

    let result = match value {
        Value::Number(number) => number.as_f64(),
        Value::String(text) => text.parse::<f64>().ok(),
        Value::Bool(flag) => Some(if *flag { 1.0 } else { 0.0 }),
        _ => None,
    };

    let Some(value) = result else {
        return SDEP_INTERNAL_ERROR;
    };

    unsafe {
        *out = value;
    }

    SDEP_SUCCESS
}

pub(crate) fn lookup_string(
    database: &str,
    field: &str,
    id: *const c_char,
    buffer: *mut u16,
    buffer_len: usize,
) -> i32 {
    if buffer.is_null() || buffer_len == 0 {
        return SDEP_INVALID_ARGUMENT;
    }

    let store = match ensure_store() {
        Ok(store) => store,
        Err(code) => {
            write_wide_string(buffer, buffer_len, "");
            return code;
        }
    };

    let Some(id) = c_string_arg(id) else {
        write_wide_string(buffer, buffer_len, "");
        return SDEP_BAD_ID;
    };

    let Some(record) = lookup_record(&store, database, &id) else {
        write_wide_string(buffer, buffer_len, "");
        return SDEP_BAD_ID;
    };
    let Some(value) = record.fields.get(field) else {
        write_wide_string(buffer, buffer_len, "");
        return SDEP_INTERNAL_ERROR;
    };

    let value = value_to_string(value).unwrap_or_default();
    write_wide_string(buffer, buffer_len, &value)
}

pub(crate) fn lookup_ids(database: &str, buffer: *mut u16, buffer_len: usize) -> i32 {
    if buffer.is_null() || buffer_len == 0 {
        return SDEP_INVALID_ARGUMENT;
    }

    let store = match ensure_store() {
        Ok(store) => store,
        Err(code) => {
            write_wide_string(buffer, buffer_len, "");
            return code;
        }
    };

    let ids = store
        .datasets
        .get(database)
        .map(|dataset| dataset.ids.join(","))
        .unwrap_or_default();

    write_wide_string(buffer, buffer_len, &ids)
}

pub(crate) fn lookup_dropdown_list(
    database: &str,
    field: &str,
    buffer: *mut u16,
    buffer_len: usize,
) -> i32 {
    if buffer.is_null() || buffer_len == 0 {
        return SDEP_INVALID_ARGUMENT;
    }

    let store = match ensure_store() {
        Ok(store) => store,
        Err(code) => {
            write_wide_string(buffer, buffer_len, "");
            return code;
        }
    };

    let Some(dataset) = store.datasets.get(database) else {
        write_wide_string(buffer, buffer_len, "");
        return SDEP_SUCCESS;
    };

    let mut entries = Vec::with_capacity(dataset.ids.len());

    for id in &dataset.ids {
        let Some(record) = dataset.records_by_id.get(id) else {
            write_wide_string(buffer, buffer_len, "");
            return SDEP_INTERNAL_ERROR;
        };
        let Some(value) = record.fields.get(field) else {
            write_wide_string(buffer, buffer_len, "");
            return SDEP_INTERNAL_ERROR;
        };
        let Some(text) = value_to_string(value) else {
            write_wide_string(buffer, buffer_len, "");
            return SDEP_INTERNAL_ERROR;
        };

        entries.push(format!("{id}\t{text}"));
    }

    write_wide_string(buffer, buffer_len, &entries.join("\n"))
}

fn c_string_arg(ptr: *const c_char) -> Option<String> {
    if ptr.is_null() {
        return None;
    }

    let value = unsafe { std::ffi::CStr::from_ptr(ptr) };
    value.to_str().ok().map(|s| s.to_owned())
}

fn write_wide_string(buffer: *mut u16, buffer_len: usize, value: &str) -> i32 {
    if buffer.is_null() || buffer_len == 0 {
        return SDEP_INVALID_ARGUMENT;
    }

    let utf16: Vec<u16> = value.encode_utf16().collect();
    let max_copy = buffer_len.saturating_sub(1);
    let copy_len = utf16.len().min(max_copy);

    unsafe {
        std::ptr::copy_nonoverlapping(utf16.as_ptr(), buffer, copy_len);
        *buffer.add(copy_len) = 0;
    }

    if utf16.len() > max_copy {
        SDEP_BUFFER_TOO_SMALL
    } else {
        SDEP_SUCCESS
    }
}

/// Adds two `f64` values.
fn add(a: f64, b: f64) -> f64 {
    a + b
}

/// C-compatible entry point exposing [`add`] across the FFI boundary.
#[unsafe(no_mangle)]
pub extern "C" fn sdep_add(a: f64, b: f64) -> f64 {
    add(a, b)
}

#[unsafe(no_mangle)]
pub extern "C" fn sdep_string_vmp(panel_vmp: f64, panels_per_string: i32) -> f64 {
    panel_vmp * panels_per_string as f64
}

#[unsafe(no_mangle)]
pub extern "C" fn sdep_string_voc(panel_voc: f64, panels_per_string: i32) -> f64 {
    panel_voc * panels_per_string as f64
}

#[unsafe(no_mangle)]
pub extern "C" fn sdep_mppt_current(panel_imp: f64, parallel_strings: i32) -> f64 {
    panel_imp * parallel_strings as f64
}

#[unsafe(no_mangle)]
pub extern "C" fn sdep_array_dc_power(panel_watts: f64, panel_count: i32) -> f64 {
    panel_watts * panel_count as f64
}

#[unsafe(no_mangle)]
pub extern "C" fn sdep_cold_voc(
    panel_voc: f64,
    panels_per_string: i32,
    temp_coeff_pct_per_c: f64,
    design_temp_c: f64,
) -> f64 {
    let delta_c = 25.0 - design_temp_c;
    panel_voc * panels_per_string as f64 * (1.0 + (temp_coeff_pct_per_c.abs() / 100.0) * delta_c)
}

#[unsafe(no_mangle)]
    pub extern "C" fn sdep_init() -> i32 {
    let dir = dll_directory()
        .map(|path| path.join("DataJson"))
        .unwrap_or_else(default_data_dir);

    match load_store_from_dir(&dir) {
        Ok(store) => {
            set_store(store);
            SDEP_SUCCESS
        }
        Err(code) => {
            set_failure(code);
            code
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::ffi::CString;

    #[test]
    fn add_computes_sum() {
        assert_eq!(add(2.0, 3.0), 5.0);
    }

    #[test]
    fn loads_all_databases_and_records() {
        let dir = temp_test_dir("loads_all");
        write_mock_dataset(
            &dir,
            "panels",
            "Panels",
            vec![
                serde_json::json!({
                    "ID": "P000001",
                    "Manufacturer": "REC",
                    "Model": "Alpha",
                    "Vmp V": 54.9
                }),
                serde_json::json!({
                    "ID": "P000002",
                    "Manufacturer": "Qcells",
                    "Model": "Beta",
                    "Vmp V": 38.39
                }),
            ],
        );
        write_mock_dataset(
            &dir,
            "inverters",
            "Inverters",
            vec![serde_json::json!({
                "ID": "I000001",
                "Manufacturer": "Sol-Ark",
                "Model": "18K",
                "AC Output kW": 18
            })],
        );
        write_mock_dataset(
            &dir,
            "batteries",
            "Batteries",
            vec![
                serde_json::json!({
                    "ID": "B000001",
                    "Manufacturer": "Discover",
                    "Model": "Helios ESS"
                }),
                serde_json::json!({
                    "ID": "B000002",
                    "Manufacturer": "Pytes",
                    "Model": "V16"
                }),
                serde_json::json!({
                    "ID": "B000003",
                    "Manufacturer": "HomeGrid",
                    "Model": "Stack'd Series 19.2"
                }),
            ],
        );

        let store = load_store_from_dir(&dir).expect("load data");

        let expected_datasets = [
            ("panels", 2usize),
            ("inverters", 1usize),
            ("batteries", 3usize),
        ];
        assert_eq!(store.datasets.len(), expected_datasets.len());

        for (database_name, record_count) in expected_datasets {
            let dataset = store
                .datasets
                .get(database_name)
                .unwrap_or_else(|| panic!("{} dataset", database_name));

            assert_eq!(dataset.ids.len(), record_count);
            assert_eq!(dataset.records_by_id.len(), record_count);
        }

        let panels = store.datasets.get("panels").expect("panels dataset");
        assert!(panels.records_by_id.contains_key("P000001"));

        let record = panels.records_by_id.get("P000001").unwrap();
        let vmp = record.fields.get("vmp").unwrap().as_f64().unwrap();
        assert_eq!(vmp, 54.9);

        let manufacturer = record.fields.get("manufacturer").unwrap().as_str().unwrap();
        assert_eq!(manufacturer, "REC");
    }

    #[test]
    fn exported_numeric_getter_reads_panel_vmp() {
        let store = load_mock_panel_store();
        set_store(store);

        let id = CString::new("P000001").expect("id");
        let mut vmp = 0.0;
        let code = sdep_get_panels_vmp(id.as_ptr(), &mut vmp);
        assert_eq!(code, SDEP_SUCCESS);
        assert_eq!(vmp, 54.9);
    }

    #[test]
    fn exported_numeric_getter_rejects_bad_id() {
        let store = load_mock_panel_store();
        set_store(store);

        let bad_id = CString::new("P999999").expect("id");
        let mut vmp = 0.0;
        let code = sdep_get_panels_vmp(bad_id.as_ptr(), &mut vmp);
        assert_eq!(code, SDEP_BAD_ID);
    }

    #[test]
    fn exported_ids_getter_returns_panel_ids() {
        let store = load_mock_panel_store();
        set_store(store);

        let mut buffer = vec![0u16; 128];
        let code = sdep_get_panels_ids(buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_SUCCESS);

        let ids = wide_buffer_to_string(&buffer);
        assert_eq!(ids, ["P000001", "P000002"].join(","));
    }

    #[test]
    fn exported_model_list_getter_returns_dropdown_values() {
        let store = load_mock_full_store();
        set_store(store);

        let mut buffer = vec![0u16; 256];

        let code = sdep_get_panels_model_list(buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_SUCCESS);
        assert_eq!(
            wide_buffer_to_string(&buffer),
            [
                "P000001\tREC Alpha Pure-RX 460",
                "P000002\tQ.TRON BLK M-G2+ 430",
                "P000003\tHiKu6 455",
                "P000004\tSIL-460/470",
                "P000005\tVertex S+ 455",
            ]
            .join("\n")
        );

        buffer.fill(0);
        let code = sdep_get_inverters_model_list(buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_SUCCESS);
        assert_eq!(
            wide_buffer_to_string(&buffer),
            ["I000001\tSol-Ark 18K-2P-LV", "I000002\tSol-Ark 15K"].join("\n")
        );

        buffer.fill(0);
        let code = sdep_get_batteries_model_list(buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_SUCCESS);
        assert_eq!(
            wide_buffer_to_string(&buffer),
            [
                "B000001\tDiscover Helios ESS",
                "B000002\tPytes V16",
                "B000003\tStack'd Series 19.2",
                "B000004\teBoost 16",
            ]
            .join("\n")
        );
    }

    #[test]
    fn exported_string_getter_returns_panel_manufacturer() {
        let store = load_mock_panel_store();
        set_store(store);

        let id = CString::new("P000001").expect("id");
        let mut buffer = vec![0u16; 32];
        let code = sdep_get_panels_manufacturer(id.as_ptr(), buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_SUCCESS);
        let manufacturer = wide_buffer_to_string(&buffer);
        assert_eq!(manufacturer, "REC");
    }

    #[test]
    fn exported_string_getter_rejects_bad_id() {
        let store = load_mock_panel_store();
        set_store(store);

        let bad_id = CString::new("P999999").expect("id");
        let mut buffer = vec![0u16; 32];
        let code = sdep_get_panels_model(bad_id.as_ptr(), buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_BAD_ID);
        assert_eq!(buffer[0], 0);
    }

    #[test]
    fn exported_string_getter_truncates_to_buffer() {
        let store = load_mock_panel_store();
        set_store(store);

        let id = CString::new("P000001").expect("id");
        let mut buffer = vec![0u16; 4];
        let code = sdep_get_panels_model(id.as_ptr(), buffer.as_mut_ptr(), buffer.len());
        assert_eq!(code, SDEP_BUFFER_TOO_SMALL);
        assert_eq!(buffer[3], 0);
    }

    #[test]
    fn load_rejects_wrong_version() {
        let dir = temp_test_dir("wrong_version");
        let json = r#"{
  "header": { "databaseName": "Panels", "recordCount": 1, "version": 2.0 },
  "records": [
    { "ID": "P000001", "Model": "X" }
  ]
}"#;
        fs::write(dir.join("DataJson").join("panels.json"), json).expect("write json");

        let result = load_store_from_dir(&dir);
        assert_eq!(result.unwrap_err(), SDEP_WRONG_DATABASE_VERSION);
    }

    #[test]
    fn load_rejects_corrupt_record_count() {
        let dir = temp_test_dir("corrupt_count");
        let json = r#"{
  "header": { "databaseName": "Panels", "recordCount": 2, "version": 1.0 },
  "records": [
    { "ID": "P000001", "Model": "X" }
  ]
}"#;
        fs::write(dir.join("DataJson").join("panels.json"), json).expect("write json");

        let result = load_store_from_dir(&dir);
        assert_eq!(result.unwrap_err(), SDEP_DATABASE_CORRUPT);
    }

    fn wide_buffer_to_string(buffer: &[u16]) -> String {
        let len = buffer.iter().position(|&c| c == 0).unwrap_or(buffer.len());
        String::from_utf16(&buffer[..len]).expect("valid utf16")
    }

    fn mock_dataset(database_name: &str, records: Vec<Value>) -> Value {
        serde_json::json!({
            "header": {
                "databaseName": database_name,
                "recordCount": records.len(),
                "version": 1.0
            },
            "records": records,
        })
    }

    fn write_mock_dataset(dir: &Path, file_stem: &str, database_name: &str, records: Vec<Value>) {
        let path = dir.join("DataJson").join(format!("{file_stem}.json"));
        let json = mock_dataset(database_name, records);
        let text = serde_json::to_string_pretty(&json).expect("serialize mock json");
        fs::write(path, text).expect("write mock dataset");
    }

    fn load_mock_panel_store() -> DatabaseStore {
        let dir = temp_test_dir("mock_panels");
        write_mock_dataset(
            &dir,
            "panels",
            "Panels",
            vec![
                serde_json::json!({
                    "ID": "P000001",
                    "Manufacturer": "REC",
                    "Model": "Alpha Pure-RX 460",
                    "Vmp V": 54.9
                }),
                serde_json::json!({
                    "ID": "P000002",
                    "Manufacturer": "Qcells",
                    "Model": "Q.TRON BLK M-G2+ 430",
                    "Vmp V": 38.39
                }),
            ],
        );
        load_store_from_dir(&dir).expect("load mock panels")
    }

    fn load_mock_full_store() -> DatabaseStore {
        let dir = temp_test_dir("mock_full");
        write_mock_dataset(
            &dir,
            "panels",
            "Panels",
            vec![
                serde_json::json!({
                    "ID": "P000001",
                    "Manufacturer": "REC",
                    "Model": "REC Alpha Pure-RX 460",
                    "Vmp V": 54.9
                }),
                serde_json::json!({
                    "ID": "P000002",
                    "Manufacturer": "Qcells",
                    "Model": "Q.TRON BLK M-G2+ 430",
                    "Vmp V": 38.39
                }),
                serde_json::json!({
                    "ID": "P000003",
                    "Manufacturer": "Canadian Solar",
                    "Model": "HiKu6 455",
                    "Vmp V": 34.6
                }),
                serde_json::json!({
                    "ID": "P000004",
                    "Manufacturer": "Silfab",
                    "Model": "SIL-460/470",
                    "Vmp V": 41.7
                }),
                serde_json::json!({
                    "ID": "P000005",
                    "Manufacturer": "Trina",
                    "Model": "Vertex S+ 455",
                    "Vmp V": 43.1
                }),
            ],
        );
        write_mock_dataset(
            &dir,
            "inverters",
            "Inverters",
            vec![
                serde_json::json!({
                    "ID": "I000001",
                    "Manufacturer": "Sol-Ark",
                    "Model": "Sol-Ark 18K-2P-LV",
                    "AC Output kW": 18
                }),
                serde_json::json!({
                    "ID": "I000002",
                    "Manufacturer": "Sol-Ark",
                    "Model": "Sol-Ark 15K",
                    "AC Output kW": 15
                }),
            ],
        );
        write_mock_dataset(
            &dir,
            "batteries",
            "Batteries",
            vec![
                serde_json::json!({
                    "ID": "B000001",
                    "Manufacturer": "Discover",
                    "Model": "Discover Helios ESS"
                }),
                serde_json::json!({
                    "ID": "B000002",
                    "Manufacturer": "Pytes",
                    "Model": "Pytes V16"
                }),
                serde_json::json!({
                    "ID": "B000003",
                    "Manufacturer": "HomeGrid",
                    "Model": "Stack'd Series 19.2"
                }),
                serde_json::json!({
                    "ID": "B000004",
                    "Manufacturer": "FranklinWH",
                    "Model": "eBoost 16"
                }),
            ],
        );
        load_store_from_dir(&dir).expect("load mock full store")
    }

    fn temp_test_dir(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!(
            "sdep_engine_{}_{}",
            name,
            std::process::id()
        ));
        let _ = fs::remove_dir_all(&dir);
        fs::create_dir_all(dir.join("DataJson")).expect("create temp dir");
        dir
    }
}

include!(concat!(env!("OUT_DIR"), "/sdep_api.rs"));
