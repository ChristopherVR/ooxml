# Collaboration

File > Share starts a live session between windows of one browser. Sharing across devices is not available yet.

::: warning Partial, same-browser only
There is no hosted service and no server transport in this build. Nothing leaves your browser, and no document is uploaded.
:::

## What exists today

- A live session built on `ooxml-core/collab` (Yjs) over a `BroadcastChannel`, so windows and tabs of the same browser can share a drawing.
- The whole VSDX package is shared, and the collab runtime loads only when sharing starts.
- Remote changes arrive as one undoable history step, and collaborators appear through the shared `ooxml-ui` presence element.
- The session name is 1 to 64 letters, digits, hyphens or underscores. A window joining a session whose room already holds a drawing adopts it; otherwise it seeds the room.

## Live demo: two windows

Both panes below are separate instances of the playground in one browser. In each, choose **File**, then **Share**, enter the same session name and start sharing. Then edit shape text in one pane and watch the other follow; undo reverts the remote step.

<iframe
	src="/visio-viewer/demo/?embed=1&sample=1"
	title="visio-viewer window A"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/visio-viewer/demo/?embed=1&sample=1"
	title="visio-viewer window B"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

You can do the same in two real tabs: [open the playground](/demo/){target="_self"} twice.

## What is and is not supported

| Supported                                                      | Not supported                                                                 |
| -------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| Same-browser sessions over BroadcastChannel, built on Yjs      | Sharing across devices: needs a host-provided server transport, none ships    |
| Whole-package sharing with one undoable step per remote change | Field-level merging: concurrent changes are not merged, the last package wins |
| Collaborators shown with the shared presence element           | Comments, authentication and roles                                            |
| File > Share as the entry point                                | A documented embedding API for custom transports                              |

The evidence is recorded in the [capability ledger](/parity) under "Collaboration and comments".
