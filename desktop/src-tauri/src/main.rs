#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod disk;

fn main() {
    tauri::Builder::default()
        .manage(disk::DiskAccess::default())
        .invoke_handler(tauri::generate_handler![
            disk::disk_choose_folder,
            disk::disk_scan,
            disk::disk_read,
            disk::disk_reveal
        ])
        .run(tauri::generate_context!())
        .expect("failed to start OOXML Office");
}
