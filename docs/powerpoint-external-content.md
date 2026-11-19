# Linked external content: what Office does, and what OOXML does

A presentation, document, workbook or drawing can link a picture or media file at a web
address instead of embedding it. Opening such a file can tell that server who opened it and when
(a web beacon). This page records what the Microsoft Office applications do with that content by
default, how we established it, and the default OOXML follows for each product.

The maintainer's decision (2026-11-19) is that PowerPoint content follows Microsoft PowerPoint's
default. Word, Excel and Visio are compared with their own applications below.

Everything here comes from public Microsoft documentation and from the observable behaviour of
installed Office through its documented COM automation. No Office binary was inspected.

## Summary

| Content                                  | Microsoft PowerPoint (default)                                                                                                                                      | OOXML PowerPoint (2.0 default)                                                                |
| ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| Linked picture on an Internet host       | Not downloaded. A placeholder says "To help protect your privacy, PowerPoint has blocked automatic download of this picture"; the Message Bar offers Enable Content | Not downloaded. A translated security bar offers Enable Content for that presentation         |
| Linked picture on an intranet host       | Downloaded when drawn                                                                                                                                               | Not downloaded until Enable Content (a browser cannot tell an intranet host apart; see below) |
| Linked audio and video                   | Not downloaded on open                                                                                                                                              | Not downloaded until Enable Content                                                           |
| Online video (YouTube, Vimeo)            | The provider's page loads when the video is played; needs the connected experience that downloads online content (on by default)                                    | Not loaded until Enable Content; after it, the provider's frame loads when the slide is shown |
| Linked OLE object                        | Not updated on open; a security notice asks to update links                                                                                                         | Never fetched; the stored preview picture is drawn                                            |
| Hyperlink to a media file or page        | Opened only when clicked                                                                                                                                            | Opened only when clicked                                                                      |
| File from the Internet (Mark of the Web) | Opens in Protected View                                                                                                                                             | Same as any file: blocked until Enable Content                                                |
| Enable Content                           | For that presentation, until it is closed; asked again on the next open                                                                                             | For that presentation, until another one is opened; never stored                              |
| Administrator override                   | Policy "Unblock automatic download of linked images" (PowerPoint Options > Security)                                                                                | Host option `allowExternalContent: true` (download without asking) or `false` (block, no bar) |
| Global user setting                      | None: "You cannot change external content settings globally for Microsoft PowerPoint in the Trust Center"                                                           | File > Options > Trust Center "Allow external content", off by default                        |

## Sources (retrieved 2026-11-19)

1. Microsoft Support, [Block or unblock external content in Office documents](https://support.microsoft.com/en-us/topic/10204ae0-0621-411f-b0d6-575b0847a795)
   (applies to PowerPoint for Microsoft 365 and 2016 to 2024). It defines external content as
   "any content linked from the Internet or an intranet to a workbook or presentation", names
   images and linked media as Web beacons ("When you open the file, the image is downloaded and
   information about the file is beaconed back to the external server"), says a business bar
   notifies the user that external content has been disabled and that **Enable Content**
   unblocks it, and states: "External content security settings are located in the Trust Center
   for Microsoft Excel only. You cannot change external content settings globally for Microsoft
   PowerPoint in the Trust Center." For Excel it gives the defaults "Prompt user about Data
   Connections" and "Prompt user on automatic update for Workbook Links".
2. Microsoft Learn (Office 2007 Resource Kit, last updated 2017-01-10),
   [Plan external content settings in the 2007 Office system](<https://learn.microsoft.com/en-us/previous-versions/office/office-2007-resource-kit/cc179053(v=office.12)>):
   "By default, links to external content are disabled. This includes links to data sources,
   hyperlinks to Web sites and documents, and links to images and media. When a user opens a
   document that contains links to external content, the Message Bar notifies the user that the
   links are disabled." For PowerPoint: "By default, Office PowerPoint 2007 does not display
   images that are saved on an external Web site ... Instead, a placeholder image displays and a
   warning appears in the Message Bar ... Users can allow the image to download by clicking the
   Message Bar", "Links to images on internal Web sites are not blocked", and the policy
   "Unblock automatic download of linked images" allows them. The same policy ships in the
   current PowerPoint administrative templates (PowerPoint Options > Security), where security
   baselines keep it disabled.
3. Microsoft Support, [PowerPoint 2013 or PowerPoint 2010 stops working after you repair a presentation](https://support.microsoft.com/en-us/topic/powerpoint-2013-or-powerpoint-2010-stops-working-after-you-repair-a-presentation-0e04c3d0-b3dc-c043-8c49-7dbe6b82bfd7):
   a linked picture shows "To help protect your privacy, PowerPoint has blocked automatic download
   of this picture", and a presentation with links to an OLE object raises a security message
   whose choices are **Update Links** and **Cancel**.
4. Microsoft Support, [What is Protected View?](https://support.microsoft.com/en-us/office/what-is-protected-view-d6f09ac7-e6b9-4495-8e43-2bbcdbcb6653):
   files from the Internet open in Protected View, a read-only mode, and cloud fonts are not
   downloaded until the user enables editing.
5. Microsoft Learn, [Connected experiences in Office](https://learn.microsoft.com/en-us/microsoft-365-apps/privacy/connected-experiences)
   (2026-06-12) and [Optional connected experiences](https://learn.microsoft.com/en-us/microsoft-365-apps/privacy/optional-connected-experiences)
   (2026-07-21): Insert online video is a connected experience that downloads online content;
   "when you connect to YouTube, the terms of service and privacy policy for YouTube apply".
   Microsoft's video help says an online video is stored on the website, not in the
   presentation, and needs an Internet connection to play.
6. Microsoft Support, [Refresh imported data](https://support.microsoft.com/en-us/visio/refresh-imported-data)
   and Microsoft Learn, [DataRecordset.RefreshInterval](https://learn.microsoft.com/en-us/office/vba/api/visio.datarecordset.refreshinterval):
   Visio refreshes linked data when the user chooses Refresh Data or turns on Automatic Refresh;
   "When you create a DataRecordset object, its RefreshInterval property is set to the default, 0. This setting indicates that data is not refreshed automatically."

The wording of PowerPoint's Message Bar, "SECURITY WARNING References to external pictures and
other media have been blocked. Enable Content", is reported by third-party help pages and was not
observed here (the bar is part of a window, and our observations use no window). OOXML uses that
wording.

## Observations (2026-11-19)

Microsoft 365 Apps 16.0.20527.20196 on Windows 11, no Office policies set, all files in a folder
that is not a Trusted Location. Each application was driven through documented COM automation on
an instance the probe started, with no window: `Presentations.Open(file, ReadOnly, Untitled,
WithWindow:=msoFalse)`, `Documents.Open` and `Workbooks.Open` on invisible applications,
`DisplayAlerts` off. The linked targets pointed at a local HTTP server that logged every request.
Three host names reached it: `127.0.0.1.nip.io` (public DNS for 127.0.0.1, so the Internet zone),
`127.0.0.1` and `localhost` (a dotless name, the Local intranet zone). The decks were authored by
PowerPoint itself (`Shapes.AddPicture` and `Shapes.AddMediaObject2` with `LinkToFile`), then
their link targets were pointed at the server. The same files were also given a Mark of the Web
(`Zone.Identifier`, ZoneId 3).

| Application | Linked to                       | Requests on open | Requests while drawing (`Slide.Export`, PDF export) | Picture drawn             |
| ----------- | ------------------------------- | ---------------- | --------------------------------------------------- | ------------------------- |
| PowerPoint  | `127.0.0.1.nip.io` (Internet)   | none             | none                                                | the privacy placeholder   |
| PowerPoint  | `127.0.0.1`                     | none             | none                                                | the privacy placeholder   |
| PowerPoint  | `localhost` (intranet)          | none             | the picture only                                    | the linked picture        |
| PowerPoint  | linked video and audio          | none             | none (also none when `MediaFormat` was read)        | the embedded poster frame |
| PowerPoint  | hyperlink to a media file       | none             | none                                                | not applicable            |
| Word        | `127.0.0.1.nip.io`, `localhost` | the picture      | none                                                | the linked picture        |
| Excel       | `127.0.0.1.nip.io`, `localhost` | the picture      | none                                                | the linked picture        |

The Mark of the Web made no difference to these automation opens, and
`Application.ProtectedViewWindows.Count` stayed 0: opening through `Presentations.Open` does not
use Protected View. Protected View itself, linked OLE objects (whose security notice is a modal
dialog) and online video playback (which needs a window) were not observed, to keep to automation
without windows or dialogs; their rows above come from the sources.

Every PowerPoint request carried the user agent `Mozilla/4.0 (compatible; ms-office; MSOffice 16)`.

## What OOXML does

### PowerPoint

The default matches PowerPoint's: opening a deck sends no request for its linked pictures, shape
and table fills, picture bullets, 3D model posters, audio, video or online video. When the deck
has such links (`PptxData.hasLinkedWebContent`), the viewer shows a translated security bar under
the ribbon: "SECURITY WARNING References to external pictures and other media have been blocked."
with **Enable Content** and a close button. Enable Content reopens that presentation with its
linked content (keeping unsaved changes) and lasts until another presentation is opened. Close
hides the bar and keeps the content blocked. Nothing is stored.

Differences from PowerPoint, all on the stricter side:

- A browser has no Windows security zones, so an intranet host is blocked like any other until
  Enable Content.
- A blocked linked picture is left empty instead of showing PowerPoint's privacy placeholder
  text.
- After Enable Content, an online video's frame loads when its slide is shown, where PowerPoint
  loads it on Play.

Hosts and users change it with:

- `allowExternalContent: true`: download without asking, as PowerPoint's "Unblock automatic
  download of linked images" policy does. The Trust Center option is locked on.
- `allowExternalContent: false`: block without a bar (Teams presentation previews always do). The
  Trust Center option is locked off.
- File > Options > Trust Center > "Allow external content (remote images and media)", off by
  default. PowerPoint has no such option; it stays for users and hosts that want automatic
  downloads.
- In the OOXML Office suite, Settings > Privacy > "Block linked pictures and media in
  presentations" passes `false`.

The core parser (`ooxml-core/pptx`) keeps its own default: `allowExternalImages` off, so a deck
loaded without a viewer references no linked URL.

### Word, Excel and Visio

These products already fetch nothing a file links to, so no change was needed:

- **Word**: linked pictures are not downloaded. Word downloads them on open (observed above), and
  files from the Internet open in Protected View. OOXML is stricter; showing linked pictures is a
  parity gap in [parity-checklist.md](parity-checklist.md).
- **Excel**: linked pictures are not downloaded, external workbook links and data connections are
  kept but never refreshed, and web functions are not evaluated. Excel downloads linked pictures on
  open (observed above) and prompts before data connections and workbook links (source 1). OOXML
  is stricter.
- **Visio**: data record sets and linked data are kept but never refreshed. Visio refreshes only
  when asked or on an interval the user sets (source 6), so OOXML matches its default.

[PRIVACY.md](../PRIVACY.md) lists every request the packages can make.

## Rerunning the observations

The probe kit (deck authoring through COM, the logging server, the open and export runs and the
raw results) is kept with the release evidence outside this repository. Rerun it only on a machine
where no other session is using Office: PowerPoint is a single-instance application, so the probe
refuses to start when PowerPoint is already running rather than attach to someone else's instance.
