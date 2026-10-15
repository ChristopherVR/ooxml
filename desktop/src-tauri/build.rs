fn main() {
    // Declare the app's own commands so they can be granted per window in capabilities/. Without an
    // app manifest every command is callable from every webview; with it, only the windows a
    // capability names (the main window) can invoke them.
    let attributes = tauri_build::Attributes::new().app_manifest(
        tauri_build::AppManifest::new().commands(&[
            "disk_choose_folder",
            "disk_scan",
            "disk_read",
            "disk_reveal",
        ]),
    );
    tauri_build::try_build(attributes).expect("failed to run tauri-build");
}
