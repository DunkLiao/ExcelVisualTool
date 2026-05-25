use std::{
    fs::{self, OpenOptions},
    io::Write,
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    init_file_logging();
    let _ = write_log_line("INFO", "Application starting");

    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            write_app_log,
            save_chart_png,
            save_query_xlsx
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[tauri::command]
fn write_app_log(level: String, message: String) -> Result<(), String> {
    write_log_line(&level, &message).map_err(|error| error.to_string())
}

#[tauri::command]
fn save_chart_png(file_name: String, image_bytes: Vec<u8>) -> Result<String, String> {
    if image_bytes.is_empty() {
        return Err("PNG image data is empty.".to_string());
    }

    let export_path = save_export_file(&file_name, "png", "chart.png", image_bytes)?;

    write_log_line("INFO", &format!("Exported chart PNG to {}", export_path.display()))
        .map_err(|error| error.to_string())?;
    Ok(export_path.display().to_string())
}

#[tauri::command]
fn save_query_xlsx(file_name: String, workbook_bytes: Vec<u8>) -> Result<String, String> {
    if workbook_bytes.is_empty() {
        return Err("XLSX workbook data is empty.".to_string());
    }

    let export_path = save_export_file(&file_name, "xlsx", "query-result.xlsx", workbook_bytes)?;

    write_log_line(
        "INFO",
        &format!("Exported query XLSX to {}", export_path.display()),
    )
    .map_err(|error| error.to_string())?;
    Ok(export_path.display().to_string())
}

fn init_file_logging() {
    std::panic::set_hook(Box::new(|panic_info| {
        let location = panic_info
            .location()
            .map(|location| format!("{}:{}", location.file(), location.line()))
            .unwrap_or_else(|| "unknown location".to_string());
        let payload = panic_info
            .payload()
            .downcast_ref::<&str>()
            .copied()
            .or_else(|| panic_info.payload().downcast_ref::<String>().map(String::as_str))
            .unwrap_or("unknown panic");

        let _ = write_log_line("PANIC", &format!("{payload} at {location}"));
    }));
}

fn write_log_line(level: &str, message: &str) -> std::io::Result<()> {
    let log_path = app_log_path();

    if let Some(log_dir) = log_path.parent() {
        fs::create_dir_all(log_dir)?;
    }

    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs())
        .unwrap_or_default();
    let sanitized_level = level.trim().to_uppercase();
    let sanitized_message = message.replace(['\r', '\n'], " ");
    let mut file = OpenOptions::new().create(true).append(true).open(log_path)?;

    writeln!(
        file,
        "[{timestamp}] [{}] {sanitized_message}",
        if sanitized_level.is_empty() {
            "INFO"
        } else {
            &sanitized_level
        }
    )
}

fn app_log_path() -> PathBuf {
    app_base_dir().join("logs").join("app.log")
}

fn app_base_dir() -> PathBuf {
    std::env::current_exe()
        .ok()
        .and_then(|exe_path| exe_path.parent().map(|parent| parent.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
}

fn save_export_file(
    file_name: &str,
    extension: &str,
    default_file_name: &str,
    file_bytes: Vec<u8>,
) -> Result<PathBuf, String> {
    let export_dir = app_base_dir().join("exports");
    fs::create_dir_all(&export_dir).map_err(|error| error.to_string())?;

    let safe_file_name = sanitize_file_name(file_name, extension, default_file_name);
    let export_path = unique_export_path(export_dir, &safe_file_name, extension);
    fs::write(&export_path, file_bytes).map_err(|error| error.to_string())?;

    Ok(export_path)
}

fn sanitize_file_name(file_name: &str, extension: &str, default_file_name: &str) -> String {
    let sanitized = file_name
        .chars()
        .map(|character| match character {
            '<' | '>' | ':' | '"' | '/' | '\\' | '|' | '?' | '*' | '\0'..='\u{1f}' => '_',
            _ => character,
        })
        .collect::<String>()
        .trim()
        .trim_matches('.')
        .to_string();
    let expected_suffix = format!(".{}", extension.trim_start_matches('.').to_lowercase());

    if sanitized.is_empty() {
        default_file_name.to_string()
    } else if sanitized.to_lowercase().ends_with(&expected_suffix) {
        sanitized
    } else {
        format!("{sanitized}{expected_suffix}")
    }
}

fn unique_export_path(export_dir: PathBuf, file_name: &str, extension: &str) -> PathBuf {
    let first_path = export_dir.join(file_name);
    if !first_path.exists() {
        return first_path;
    }

    let suffix = format!(".{}", extension.trim_start_matches('.'));
    let stem = file_name.strip_suffix(&suffix).unwrap_or(file_name);
    for index in 1.. {
        let candidate = export_dir.join(format!("{stem}-{index}{suffix}"));
        if !candidate.exists() {
            return candidate;
        }
    }

    unreachable!("unbounded filename suffix search should always return");
}
