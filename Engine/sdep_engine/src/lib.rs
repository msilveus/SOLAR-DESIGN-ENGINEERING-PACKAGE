/// Adds two `f64` values.
fn add(a: f64, b: f64) -> f64 {
    a + b
}

/// C-compatible entry point exposing [`add`] across the FFI boundary.
#[unsafe(no_mangle)]
pub extern "C" fn sdep_add(a: f64, b: f64) -> f64 {
    add(a, b)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn add_computes_sum() {
        assert_eq!(add(2.0, 3.0), 5.0);
    }
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
    design_temp_c: f64
) -> f64 {
    let delta_c = 25.0 - design_temp_c;
    panel_voc * panels_per_string as f64 * (1.0 + (temp_coeff_pct_per_c.abs() / 100.0) * delta_c)
}
