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
        .invoke_handler(tauri::generate_handler![write_app_log])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[tauri::command]
fn write_app_log(level: String, message: String) -> Result<(), String> {
    write_log_line(&level, &message).map_err(|error| error.to_string())
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
    std::env::current_exe()
        .ok()
        .and_then(|exe_path| exe_path.parent().map(|parent| parent.to_path_buf()))
        .unwrap_or_else(|| std::env::current_dir().unwrap_or_else(|_| PathBuf::from(".")))
        .join("logs")
        .join("app.log")
}
