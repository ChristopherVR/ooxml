# Configuring the suite assistant

The OOXML Office assistant is governed by deployment configuration, not by its UI. An organisation decides which providers and models are offered, who holds the keys, which features are on and what every request is told; users pick among what is offered and may adjust only what the policy leaves to them.

## Where the configuration comes from

1. A host that embeds the suite (for example the desktop shell) can set `globalThis.ooxmlAiConfig` to a configuration object before the app loads.
2. Otherwise the app reads `ai-config.json` beside `index.html` (the build copies `site/ai-config.json` to `suite-dist/`). It is fetched without caching and never precached by the service worker, so a changed file applies on the next load.
3. Without either, the defaults apply.

The shipped `site/ai-config.json` carries `"managed": false` and the default values, so out of the box users see the offline preview and may add their own models. Any file without `"managed": false` marks the workspace as managed: the assistant settings say that the organisation manages them.

Invalid values are ignored field by field, falling back to the defaults; a connection whose provider is unknown is dropped.

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

| Setting                  | Default                       | Meaning                                                                                                                                                 |
| ------------------------ | ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `connections`            | `[]`                          | Models the deployment provides (see below).                                                                                                             |
| `defaultConnection`      | first                         | The id selected for a user who has not chosen (`"offline"` or a connection id).                                                                         |
| `allowUserConnections`   | `true`                        | Whether users may add their own models and keys. When `false`, the model menu lists only the offline preview (if on) and `connections`.                 |
| `userProviders`          | all                           | The provider ids users may add, when they may add any.                                                                                                  |
| `offline.enabled`        | `true`                        | Offers the offline preview, which answers from rules on the device and drafts sample files without a model.                                             |
| `features.editFiles`     | `true`                        | Lets replies propose edits to attached files. When off, the model is told it cannot, and edit proposals are rejected.                                   |
| `features.createFiles`   | `true`                        | Lets replies propose new Word, Excel and PowerPoint files.                                                                                              |
| `features.attachFiles`   | `true`                        | Lets users attach files (Attach, `@` mentions, the open document). When off, no file content is ever sent.                                              |
| `features.teamsContext`  | `true`                        | Lets users include the current Teams channel.                                                                                                           |
| `features.history`       | `true`                        | Keeps conversations on the device. When off, chats exist only until the page closes and the history list is hidden.                                     |
| `features.versions`      | `true`                        | Keeps the previous bytes of a file as a version on every save and assistant edit (at most 30 per file).                                                 |
| `limits.maxContextChars` | `150000`                      | The largest attached context, as JSON characters, a request may send.                                                                                   |
| `limits.historyMessages` | `20`                          | Earlier messages sent with each request; `0` sends none.                                                                                                |
| `limits.maxOutputTokens` | `8192`                        | The reply length requested from the provider.                                                                                                           |
| `behaviour.instructions` | `""`                          | Added to the suite's system prompt for every request, after the suite's own rules.                                                                      |
| `behaviour.style`        | `balanced`                    | `concise`, `balanced` or `detailed`.                                                                                                                    |
| `behaviour.language`     | `""`                          | The language replies are written in; empty follows the user.                                                                                            |
| `behaviour.temperature`  | unset                         | Sent to the provider when set (0 to 2).                                                                                                                 |
| `userControls`           | style, language, instructions | Which of `style`, `language`, `instructions` and `temperature` a user may set for themselves. A user's instructions are added after the organisation's. |
| `notice`                 | `""`                          | A line shown in the assistant settings, for example who to ask for access.                                                                              |

Whatever the policy, a user can still turn three conveniences on or off for themselves: attaching the open document when the assistant opens beside it, opening new files as soon as they are created, and showing replies as they are written.

## Connections

| Field      | Meaning                                                                                                                                                  |
| ---------- | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `id`       | Letters, digits, `-` and `_`; unique.                                                                                                                    |
| `label`    | The name in the model menu.                                                                                                                              |
| `provider` | One of the provider ids below.                                                                                                                           |
| `endpoint` | The chat endpoint; defaults to the provider's.                                                                                                           |
| `model`    | The model name; or `models`, a list the user chooses from in the model menu.                                                                             |
| `keyMode`  | `user` (each user enters their own key, kept in the tab unless they choose to remember it) or `none` (no key in the browser: a gateway or local server). |
| `headers`  | Extra request headers, for example a gateway's tenant id. Never put secrets here: the file is public to everyone who can load the app.                   |

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
