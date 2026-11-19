# Configuring the suite assistant

The OOXML Office assistant is governed by deployment configuration, not by its UI. An organisation decides which providers and models are offered, who holds the keys, which features are on and what every request is told; users pick among what is offered and may adjust only what the policy leaves to them.

## Where the configuration comes from

1. A host that embeds the suite (for example the desktop shell) can set `globalThis.ooxmlAiConfig` to a configuration object before the app loads.
2. Otherwise the app reads `ai-config.json` beside `index.html` (the build copies `site/ai-config.json` to `suite-dist/`). It is fetched without caching and never precached by the service worker, so a changed file applies on the next load.
3. Without either, **the assistant is off**: no chat page, no Ask AI, no assistant pane, and the suite opens on its files. Any configuration turns it on unless it says `"enabled": false`.

The `ooxml-office` package and the standalone apps therefore ship without an assistant. The suite demo (`site/`, published to GitHub Pages and served by `bun run preview:suite`) provides `site/ai-config.json`, which turns it on with `"managed": false` and the default values, so the demo shows the offline preview and lets users add their own models. The desktop shell consumes the same build output, so it shows the assistant too unless it sets `globalThis.ooxmlAiConfig = { enabled: false }`. Any file without `"managed": false` marks the workspace as managed: the assistant settings say that the organisation manages them.

Invalid values are ignored field by field, falling back to the defaults; a connection whose provider is unknown is dropped. The suite only loads the configuration: `normalizeAiPolicy` and `userConnectionRefusal` in `ooxml-core/ai` normalise it and decide what a user may add, so another host applies the same rules. The policy runs in the browser and guides an honest client; rules that must hold against a modified client belong in a gateway.

## Example

```json
{
	"connections": [
		{
			"id": "team-claude",
			"label": "Claude (company)",
			"provider": "gateway",
			"endpoint": "https://ai.example.com/v1/chat/completions",
			"models": ["claude-sonnet-5-5", "claude-haiku-5-5"],
			"keyMode": "none"
		},
		{
			"id": "github",
			"label": "GitHub Models",
			"provider": "github",
			"model": "openai/gpt-4.1",
			"keyMode": "user"
		}
	],
	"defaultConnection": "team-claude",
	"allowUserConnections": false,
	"offline": { "enabled": false },
	"features": { "teamsContext": false },
	"limits": { "maxContextChars": 100000, "historyMessages": 10, "maxOutputTokens": 4096 },
	"behaviour": {
		"instructions": "Use British spelling. Never include customer names in new files.",
		"style": "concise",
		"temperature": 0.3
	},
	"userControls": { "style": false, "language": true, "instructions": false },
	"notice": "Ask the IT service desk for access to more models."
}
```

## Settings

| Setting                        | Default                                   | Meaning                                                                                                                                                 |
| ------------------------------ | ----------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `enabled`                      | `false` without a config, `true` with one | Whether the assistant exists at all.                                                                                                                    |
| `connections`                  | `[]`                                      | Models the deployment provides (see below).                                                                                                             |
| `defaultConnection`            | first                                     | The id selected for a user who has not chosen (`"offline"` or a connection id).                                                                         |
| `allowUserConnections`         | `true`                                    | Whether users may add their own models and keys. When `false`, the model menu lists only the offline preview (if on) and `connections`.                 |
| `userProviders`                | all                                       | The provider ids users may add, when they may add any. Models a user saved earlier are offered only while their provider is still listed.               |
| `offline.enabled`              | `true`                                    | Offers the offline preview, which answers from rules on the device and drafts sample files without a model.                                             |
| `features.editFiles`           | `true`                                    | Lets replies propose edits to attached files. When off, the model is told it cannot, and edit proposals are rejected.                                   |
| `features.createFiles`         | `true`                                    | Lets replies propose new Word, Excel and PowerPoint files.                                                                                              |
| `features.attachFiles`         | `true`                                    | Lets users attach files (Attach, `@` mentions, the open document). When off, no file content is ever sent.                                              |
| `features.teamsContext`        | `true`                                    | Lets users include the current Teams channel.                                                                                                           |
| `features.history`             | `true`                                    | Keeps conversations on the device. When off, chats exist only until the page closes and the history list is hidden.                                     |
| `features.versions`            | `true`                                    | Keeps the previous bytes of a file as a version on every save and assistant edit (at most 30 per file).                                                 |
| `features.sharedChats`         | `true`                                    | Lets users share a chat and join chats shared with them (see below).                                                                                    |
| `collaboration.syncUrl`        | `""`                                      | The y-websocket server shared chats sync through (`ws://` or `wss://`).                                                                                 |
| `collaboration.credentialsUrl` | `""`                                      | A host endpoint (`https:` or the app's own origin) that issues the sync server's short-lived token; see [Shared chats](#shared-chats). Needs `syncUrl`. |
| `limits.maxContextChars`       | `150000`                                  | The largest attached context, as JSON characters, a request may send.                                                                                   |
| `limits.historyMessages`       | `20`                                      | Earlier messages sent with each request; `0` sends none.                                                                                                |
| `limits.maxOutputTokens`       | `8192`                                    | The reply length requested from the provider.                                                                                                           |
| `context.includeHidden`        | `false`                                   | Sends hidden and very hidden sheets and the cells of hidden rows and columns with attached workbooks. When off, they are left out and counted.          |
| `behaviour.instructions`       | `""`                                      | Added to the suite's system prompt for every request, after the suite's own rules.                                                                      |
| `behaviour.style`              | `balanced`                                | `concise`, `balanced` or `detailed`.                                                                                                                    |
| `behaviour.language`           | `""`                                      | The language replies are written in; empty follows the user.                                                                                            |
| `behaviour.temperature`        | unset                                     | Sent to the provider when set (0 to 2).                                                                                                                 |
| `userControls`                 | style, language, instructions             | Which of `style`, `language`, `instructions` and `temperature` a user may set for themselves. A user's instructions are added after the organisation's. |
| `notice`                       | `""`                                      | A line shown in the assistant settings, for example who to ask for access.                                                                              |

Whatever the policy, a user can still turn three conveniences on or off for themselves: attaching the open document when the assistant opens beside it, opening new files as soon as they are created, and showing replies as they are written.

## Connections

| Field      | Meaning                                                                                                                                                             |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `id`       | Letters, digits, `-` and `_`; unique.                                                                                                                               |
| `label`    | The name in the model menu.                                                                                                                                         |
| `provider` | One of the provider ids below.                                                                                                                                      |
| `endpoint` | The chat endpoint; defaults to the provider's.                                                                                                                      |
| `model`    | The model name; or `models`, a list the user chooses from in the model menu.                                                                                        |
| `keyMode`  | `user` (each user enters their own key, kept in the tab's session storage and cleared when it closes) or `none` (no key in the browser: a gateway or local server). |
| `headers`  | Extra request headers, for example a gateway's tenant id. Never put secrets here: the file is public to everyone who can load the app.                              |

Keys never belong in `ai-config.json`. To give users a model without handing them a key, run an OpenAI-compatible gateway that authenticates them (for example with your single sign-on, through cookies it scopes itself) and adds the provider key on the server, and configure it with `"provider": "gateway"` and `"keyMode": "none"`.

## Providers

Requests go straight from the browser to the endpoint, so the provider must allow cross-origin requests from the app's origin; providers that do not are reached through a gateway.

| Id           | Provider                         | Protocol                   | Key                      |
| ------------ | -------------------------------- | -------------------------- | ------------------------ |
| `anthropic`  | Anthropic                        | Messages API               | `x-api-key`              |
| `openai`     | OpenAI                           | Chat completions           | Bearer                   |
| `github`     | GitHub Models                    | Chat completions           | GitHub token, Bearer     |
| `azure`      | Azure OpenAI                     | Chat completions           | `api-key`                |
| `google`     | Google Gemini                    | OpenAI-compatible endpoint | Bearer                   |
| `mistral`    | Mistral                          | Chat completions           | Bearer                   |
| `groq`       | Groq                             | Chat completions           | Bearer                   |
| `deepseek`   | DeepSeek                         | Chat completions           | Bearer                   |
| `xai`        | xAI                              | Chat completions           | Bearer                   |
| `together`   | Together AI                      | Chat completions           | Bearer                   |
| `fireworks`  | Fireworks AI                     | Chat completions           | Bearer                   |
| `perplexity` | Perplexity                       | Chat completions           | Bearer                   |
| `openrouter` | OpenRouter                       | Chat completions           | Bearer                   |
| `ollama`     | Ollama on the user's computer    | Chat completions (local)   | none                     |
| `lmstudio`   | LM Studio on the user's computer | Chat completions (local)   | none                     |
| `gateway`    | An organisation gateway          | Chat completions           | none (added server-side) |
| `custom`     | Any OpenAI-compatible endpoint   | Chat completions           | Bearer                   |

Providers without a browser-callable API (for example services that need request signing) are reached through a gateway.

## Shared chats

A shared chat is a conversation that lives in its own collaboration room (`createSharedChat` in `ooxml-core/teams`), so people in other workspaces, profiles and sessions work in the same chat. **Share** on the chat page moves the conversation into a new room and shows an invite link (`#/chat/join/aichat-<32 hex>`); **Join** in the history opens a pasted link.

- The room syncs through `collaboration.syncUrl` when the policy sets one, otherwise through the sync server of the Teams connection the user configured, otherwise only between the tabs of one browser (BroadcastChannel). The share dialog says which applies.
- The configuration is public to everyone who can load the app, so it holds no token. A sync server that needs one gets it from `collaboration.credentialsUrl`: right before a shared chat connects, and again before every reconnect or once the token expires, the suite sends a `GET` to that URL with the user's cookies (`credentials: 'include'`) and expects `{ "token": "...", "expiresAt": "<ISO 8601 time>" }`. Authenticate the user there (for example with your single sign-on) and issue a token that expires within minutes. The token is held in memory only, never in storage, and travels in the socket's `Sec-WebSocket-Protocol` header. If the request fails, the chat keeps working on this device and the bar above the composer says why it is not syncing. A `collaboration.token` in older configurations is ignored.
- Whoever asks runs the request with their own model connection and streams the reply into the room; the others watch it arrive and see who is typing or asking.
- The room carries each message's public projection only: the request text, the reply text as it is shown, the model's label, the names and kinds of the files a request attached, and how many edits and new files a reply proposed. The reply text is the model's answer, so it can quote attached files. Attached files, the reply's raw text, its proposed edits and new files, their before-and-after text, and the files' ids and revisions stay on the asker's device.
- Only the person who asked sees the proposal cards and can review, apply or create what a reply proposed; the others see a note with the counts. A message someone else wrote is display only on every other page: it never becomes a card, and when a later request sends history, it is quoted inside a user turn rather than sent as the assistant's own words.
- Each page writes only its own messages to the room. In a conversation shared before this version, every earlier message becomes display only once the room syncs, so the proposals in it can no longer be applied from the chat; ask again to get new ones.
- The room id is the only key, so treat invite links as secrets, and authorise rooms on the sync server for anything sensitive. The sync server can read everything in the room, replies included, as with Teams chat.

## Reviewing proposed edits in the document

Every proposal card offers **Review in document**, which opens the file in its editor with the suggestions in place:

- **Word** shows them as real tracked changes by "Assistant" (`suggestDocxRunText` in `ooxml-core/automation`). Accept or reject each one from the Review tab and choose **Close** to keep the result in the open document, or use **Accept all** / **Reject all** in the suggestions panel.
- **Excel, PowerPoint and Visio** have no tracked changes in their file formats, so the editor shows the file as it would be, read only (Visio has no read-only switch; edits made there while reviewing are discarded). The panel selects each change in the editor (the cells, the slide's shape, or the Visio page and shape), each one can be undone or kept, and **Apply kept changes** saves them as one version. **Discard** returns to the saved file.

While the suggestions are on screen the document is a preview, and nothing takes it out of the suite until the panel applies, discards or (Word) keeps them. The suite's Save and Download, the editor's own Save and Ctrl+P refuse with a message naming the document. The editor's commands that write, print, share or copy the document (Save As, Export and its formats, Download a copy, Print, Share, Copy and Cut, and Excel's CSV and Visio's shape report export) are switched off in the File tab, the ribbon, the menus and the shortcuts, and the panel says so. The browser's own Print leaves the document out. In PowerPoint, text selected with the browser can still be copied, and an embedded file can still be downloaded in its original form, which a review never changes. A message sent to the assistant meanwhile reads the file as last saved.

The request reads each attached file within its share of `limits.maxContextChars`: Word documents paragraph by paragraph, workbooks as every filled cell of every visible sheet, decks slide by slide and drawings page by page. Hidden and very hidden sheets (names included) and the cells of hidden rows and columns are left out unless `context.includeHidden` is on. When something is left out, to fit or because it is hidden, the view carries an `omitted` note saying what, and the model is told to say so.

## Formulas in proposals

A proposed cell edit or new workbook may use only ordinary calculation. A formula that can reach outside the workbook (`WEBSERVICE`, `FILTERXML`, `HYPERLINK`, `IMAGE`, `RTD`, `STOCKHISTORY`, the `CUBE` functions, `PY` and similar), calls an add-in function (`_xll.`), evaluates text as a formula or runs code through the Excel 4.0 macro functions (`EVALUATE`, `EXEC` and similar), sends cell data to an AI service (`COPILOT`, `TRANSLATE`, `DETECTLANGUAGE`), refers to another workbook or a file, or cannot be read by the formula parser (a DDE link) is rejected when the reply arrives, with the cell named, and the core refuses it again when the edit is applied or the file is built. The check reads the formula with the core's parser, so a function passed by name to `MAP`, `LET` or `LAMBDA` or written with a storage prefix is found too, and text inside a string literal is not. There is no policy switch for it; hosts that build their own tooling choose with `formulas` in `ooxml-core/automation` (see [mcp.md](mcp.md)).
