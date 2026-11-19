#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod disk;
mod navigation;

use tauri::webview::NewWindowResponse;

fn main() {
    tauri::Builder::default()
        .manage(disk::DiskAccess::default())
        .invoke_handler(tauri::generate_handler![
            disk::disk_choose_folder,
            disk::disk_scan,
            disk::disk_read,
            disk::disk_reveal
        ])
        .setup(|app| {
            // The window is declared in tauri.conf.json with `create: false`, so it is built here
            // with the navigation and new-window rules attached before it loads anything.
            let config = app
                .config()
                .app
                .windows
                .iter()
                .find(|window| window.label == "main")
                .expect("tauri.conf.json declares the main window")
                .clone();
            tauri::WebviewWindowBuilder::from_config(app.handle(), &config)?
                .on_navigation(navigation::allow_navigation)
                .on_new_window(|url, _features| {
                    navigation::open_in_system(&url);
                    NewWindowResponse::Deny
                })
                .build()?;
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("failed to start OOXML Office");
}
