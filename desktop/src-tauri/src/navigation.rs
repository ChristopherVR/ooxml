//! Where the main window may go. The window only ever shows the bundled app: a link to the web
//! opens in the system browser, and nothing else (another site, a file, a blob or data document,
//! a new in-app window) loads inside the app, where it would sit next to the IPC bridge.
use tauri::Url;

#[derive(Debug, PartialEq, Eq)]
pub enum Destination {
    /// The bundled app itself.
    App,
    /// A web or mail link, handed to the system's default handler.
    System,
    /// Anything else, which is refused.
    Refused,
}

/// The origins the bundled app is served from: `tauri://localhost` on macOS and Linux,
/// `http(s)://tauri.localhost` on Windows.
fn is_app(url: &Url) -> bool {
    match url.scheme() {
        "tauri" => url.host_str() == Some("localhost"),
        "http" | "https" => url.host_str() == Some("tauri.localhost") && url.port().is_none(),
        _ => false,
    }
}

pub fn classify(url: &Url) -> Destination {
    if is_app(url) {
        return Destination::App;
    }
    match url.scheme() {
        "http" | "https" if url.username().is_empty() && url.password().is_none() => {
            Destination::System
        }
        "mailto" => Destination::System,
        _ => Destination::Refused,
    }
}

/// Opens a link classified as [`Destination::System`] in the default browser or mail client. The
/// URL goes to the handler as one argument, never through a shell.
pub fn open_in_system(url: &Url) {
    if classify(url) != Destination::System {
        return;
    }
    let target = url.as_str();
    #[cfg(target_os = "windows")]
    let spawned = {
        use std::os::windows::process::CommandExt;
        std::process::Command::new("rundll32.exe")
            .arg("url.dll,FileProtocolHandler")
            .arg(target)
            .creation_flags(0x08000000)
            .spawn()
    };
    #[cfg(target_os = "macos")]
    let spawned = std::process::Command::new("open").arg(target).spawn();
    #[cfg(all(unix, not(target_os = "macos")))]
    let spawned = std::process::Command::new("xdg-open").arg(target).spawn();
    if let Err(error) = spawned {
        eprintln!("Could not open {target} in the system browser: {error}");
    }
}

/// The main window's navigation rule: stay on the app, send links out, refuse the rest.
pub fn allow_navigation(url: &Url) -> bool {
    match classify(url) {
        Destination::App => true,
        Destination::System => {
            open_in_system(url);
            false
        }
        Destination::Refused => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn at(url: &str) -> Destination {
        classify(&Url::parse(url).unwrap())
    }

    #[test]
    fn the_bundled_app_stays_in_the_window() {
        assert_eq!(at("tauri://localhost/index.html?suite"), Destination::App);
        assert_eq!(at("http://tauri.localhost/"), Destination::App);
        assert_eq!(at("https://tauri.localhost/apps/word/"), Destination::App);
    }

    #[test]
    fn web_and_mail_links_go_to_the_system() {
        assert_eq!(at("https://example.com/page"), Destination::System);
        assert_eq!(at("http://example.com/"), Destination::System);
        assert_eq!(at("mailto:someone@example.com"), Destination::System);
    }

    #[test]
    fn look_alike_and_local_documents_are_refused_or_leave() {
        assert_eq!(at("http://tauri.localhost:8080/"), Destination::System);
        assert_eq!(at("https://tauri.localhost.example.com/"), Destination::System);
        assert_eq!(at("https://user:pass@example.com/"), Destination::Refused);
        assert_eq!(at("file:///C:/Windows/win.ini"), Destination::Refused);
        assert_eq!(at("ftp://example.com/file"), Destination::Refused);
        assert_eq!(at("data:text/html,hello"), Destination::Refused);
        assert_eq!(at("blob:http://tauri.localhost/1234"), Destination::Refused);
        assert_eq!(at("javascript:void(0)"), Destination::Refused);
    }
}
