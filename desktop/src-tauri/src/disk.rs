use serde::Serialize;
use std::{
    collections::HashSet,
    fs,
    io::Read,
    path::{Component, Path, PathBuf},
    sync::{Arc, Mutex},
};

#[derive(Default, Clone)]
pub struct DiskAccess(pub Arc<Mutex<HashSet<PathBuf>>>);
#[derive(Serialize)]
pub struct DiskEntry {
    name: String,
    path: String,
}
#[derive(Serialize)]
pub struct Scan {
    files: Vec<DiskEntry>,
    truncated: bool,
    skipped: usize,
}
fn approved(state: &DiskAccess, root: &str) -> Result<PathBuf, String> {
    let path = fs::canonicalize(root).map_err(|e| e.to_string())?;
    if !state.0.lock().map_err(|e| e.to_string())?.contains(&path) {
        return Err("Reconnect this folder using the folder picker.".into());
    }
    Ok(path)
}
fn resolve_file(root: &Path, relative: &str) -> Result<PathBuf, String> {
    let path = Path::new(relative);
    if path
        .components()
        .any(|c| !matches!(c, Component::Normal(_)))
    {
        return Err("Choose a file inside the connected folder.".into());
    }
    let file = fs::canonicalize(root.join(path)).map_err(|e| e.to_string())?;
    if !file.starts_with(root) || !file.is_file() {
        return Err("File is outside the connected folder or no longer exists.".into());
    }
    Ok(file)
}
#[tauri::command]
pub async fn disk_choose_folder(
    state: tauri::State<'_, DiskAccess>,
) -> Result<Option<String>, String> {
    let access = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let Some(path) = rfd::FileDialog::new()
            .set_title("Connect a folder to OOXML Office")
            .pick_folder()
        else {
            return Ok(None);
        };
        let path = fs::canonicalize(path).map_err(|e| e.to_string())?;
        access
            .0
            .lock()
            .map_err(|e| e.to_string())?
            .insert(path.clone());
        Ok(Some(path.to_string_lossy().into_owned()))
    })
    .await
    .map_err(|e| e.to_string())?
}
fn scan(root: PathBuf) -> Scan {
    let mut result = Scan {
        files: Vec::new(),
        truncated: false,
        skipped: 0,
    };
    let mut pending = vec![(root.clone(), 0)];
    let mut visited = 0;
    while let Some((dir, depth)) = pending.pop() {
        if depth > 40 {
            result.skipped += 1;
            continue;
        }
        let Ok(list) = fs::read_dir(dir) else {
            result.skipped += 1;
            continue;
        };
        for entry in list {
            visited += 1;
            if visited > 20000 {
                result.truncated = true;
                return result;
            }
            let Ok(entry) = entry else {
                result.skipped += 1;
                continue;
            };
            let Ok(kind) = entry.file_type() else {
                result.skipped += 1;
                continue;
            };
            if kind.is_symlink() {
                continue;
            }
            if kind.is_dir() {
                pending.push((entry.path(), depth + 1));
            } else if kind.is_file() {
                if result.files.len() >= 5000 {
                    result.truncated = true;
                    return result;
                }
                if let Ok(path) = entry.path().strip_prefix(&root) {
                    result.files.push(DiskEntry {
                        name: entry.file_name().to_string_lossy().into_owned(),
                        path: path.to_string_lossy().replace('\\', "/"),
                    });
                }
            }
        }
    }
    result
}
#[tauri::command]
pub async fn disk_scan(state: tauri::State<'_, DiskAccess>, root: String) -> Result<Scan, String> {
    let root = approved(&state, &root)?;
    tauri::async_runtime::spawn_blocking(move || scan(root))
        .await
        .map_err(|e| e.to_string())
}
#[tauri::command]
pub async fn disk_read(
    state: tauri::State<'_, DiskAccess>,
    root: String,
    path: String,
) -> Result<Vec<u8>, String> {
    let root = approved(&state, &root)?;
    tauri::async_runtime::spawn_blocking(move || {
        let path = resolve_file(&root, &path)?;
        let mut bytes = Vec::new();
        fs::File::open(path)
            .map_err(|e| e.to_string())?
            .take(33_554_433)
            .read_to_end(&mut bytes)
            .map_err(|e| e.to_string())?;
        if bytes.len() > 33_554_432 {
            return Err("Choose a file smaller than 32 MB.".into());
        }
        Ok(bytes)
    })
    .await
    .map_err(|e| e.to_string())?
}
#[tauri::command]
pub async fn disk_reveal(
    state: tauri::State<'_, DiskAccess>,
    root: String,
    path: String,
) -> Result<(), String> {
    let root = approved(&state, &root)?;
    tauri::async_runtime::spawn_blocking(move || {
        let file = resolve_file(&root, &path)?;
        #[cfg(target_os = "windows")]
        {
            use std::os::windows::process::CommandExt;
            let file = file.to_string_lossy();
            let file = if let Some(share) = file.strip_prefix(r"\\?\UNC\") {
                format!(r"\\{}", share)
            } else {
                file.strip_prefix(r"\\?\").unwrap_or(&file).to_string()
            };
            std::process::Command::new("explorer.exe")
                .arg(format!("/select,{}", file))
                .creation_flags(0x08000000)
                .spawn()
                .map_err(|e| e.to_string())?;
        }
        #[cfg(target_os = "macos")]
        {
            std::process::Command::new("open")
                .arg("-R")
                .arg(file)
                .spawn()
                .map_err(|e| e.to_string())?;
        }
        #[cfg(not(any(target_os = "windows", target_os = "macos")))]
        {
            std::process::Command::new("xdg-open")
                .arg(file.parent().ok_or("Folder is unavailable")?)
                .spawn()
                .map_err(|e| e.to_string())?;
        }
        Ok(())
    })
    .await
    .map_err(|e| e.to_string())?
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn scopes_reads_and_search_to_an_explicit_folder() {
        let dir = std::env::temp_dir().join(format!("ooxml-disk-test-{}", std::process::id()));
        fs::create_dir_all(dir.join("nested")).unwrap();
        fs::write(dir.join("nested/brief.docx"), b"test").unwrap();
        let root = fs::canonicalize(&dir).unwrap();
        let state = DiskAccess::default();
        assert!(approved(&state, root.to_str().unwrap()).is_err());
        state.0.lock().unwrap().insert(root.clone());
        assert!(approved(&state, root.to_str().unwrap()).is_ok());
        assert!(resolve_file(&root, "../outside.txt").is_err());
        assert!(resolve_file(&root, root.to_str().unwrap()).is_err());
        assert!(resolve_file(&root, "nested/brief.docx").is_ok());
        let result = scan(root);
        assert_eq!(result.files.len(), 1);
        assert_eq!(result.files[0].path, "nested/brief.docx");
        fs::remove_file(dir.join("nested/brief.docx")).unwrap();
        fs::remove_dir(dir.join("nested")).unwrap();
        fs::remove_dir(dir).unwrap();
    }
}
