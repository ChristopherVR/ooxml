# 智能助手 {#ai-assistant}

每个框架组件都提供可选的 AI 聊天面板。用户点击工具栏中的 AI 星光图标后，聊天面板会在组件右侧打开。助手可以阅读演示文稿、回答相关问题，并提出文本、样式、布局、幻灯片和主题的修改建议。默认流程中，用户可以在实际修改前审阅、接受或拒绝建议。

该功能**不绑定模型提供商**，认证也由宿主提供。组件不附带模型或 API 密钥。你可以接入自己的后端路由、浏览器内模型或完全自定义的传输层。未传入 `ai` 配置时，不显示图标，也不加载相关代码和 `ai` SDK。

## 安装 {#installation}

AI 相关包都是**可选同级依赖**。安装 `ai`，并根据 React、Vue 或 Svelte 框架安装对应的 AI SDK UI 包。Angular 和 Vanilla 使用组件内置的不依赖框架的控制器，因此只需要 `ai`。

```bash
# React
bun add ai @ai-sdk/react
# Vue
bun add ai @ai-sdk/vue
# Svelte
bun add ai @ai-sdk/svelte
# Angular / Vanilla
bun add ai
```

还需要为所选模型安装相应的**提供商**包。它由宿主自行选择，不属于预览组件的依赖：

```bash
# Any OpenAI-compatible endpoint (Ollama, vLLM, LM Studio, a proxy, ...)
bun add @ai-sdk/openai-compatible
# or a first-party provider
bun add @ai-sdk/openai
bun add @ai-sdk/anthropic
```

## 启用面板 {#enabling-the-panel}

把助手传给预览器的 `ai` 属性（Angular 中是输入，Vanilla 中是选项）。每个绑定都有一个 `/ai` 入口，其中的 `createAiAssistant(config)` 把配置转换为助手句柄。配置的结构在所有绑定中完全相同，只是传递属性的宿主语法不同。

```tsx
// React
import { PowerPointViewer } from 'pptx-react-viewer';
import { createAiAssistant } from 'pptx-react-viewer/ai';

const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });

<PowerPointViewer content={deck} ai={ai} />;
```

```ts
// Vue、Angular、Svelte 和 Vanilla：从各自绑定的 `/ai` 入口导入同一个工厂函数
import { createAiAssistant } from 'pptx-vue-viewer/ai';
// 或 'pptx-angular-viewer/ai'、'pptx-svelte-viewer/ai'、'pptx-vanilla-viewer/ai'

const ai = createAiAssistant({ connection: { kind: 'endpoint', api: '/api/ai/chat' } });
```

```vue
<!-- Vue -->
<PowerPointViewer :content="deck" :ai="ai" />
```

```html
<!-- Angular -->
<pptx-viewer [content]="deck" [ai]="ai" />
```

```svelte
<!-- Svelte -->
<PowerPointViewer source={deck} {ai} />
```

```ts
// Vanilla
import { createPptxViewer } from 'pptx-vanilla-viewer';

createPptxViewer(container, { source: deck, ai });
```

直接传入普通配置（`ai: { connection }`）在所有绑定中仍然有效，本版本也不会将其弃用。请只创建一次句柄（在模块作用域、`useMemo` 中或作为组件字段）；即使每次渲染都重新创建句柄，已打开的对话也不会被重新挂载。本版本保证什么、下一个主版本会改变什么，见 [`/ai` 入口](#ai-entry)。

## 连接方式 {#connection-modes}

`connection` 指定助手如何连接模型，共有三种方式。

### `endpoint`（推荐用于生产环境） {#endpoint-recommended-for-production}

将消息发送到自己的后端路由，模型提供商的密钥保留在服务器端，浏览器无法访问。这是更安全的默认选择。

```ts
const ai: PptxAiConfig = {
	connection: {
		kind: 'endpoint',
		api: '/api/ai/chat',
		// Headers and body may be functions, so auth tokens can be resolved per
		// request (for example from an OAuth session that refreshes).
		headers: async () => ({ Authorization: `Bearer ${await getAccessToken()}` }),
		credentials: 'include',
	},
};
```

服务器使用 AI SDK 调用模型，并以**仅包含 schema** 的形式注册演示文稿工具。工具没有 `execute`，因此模型发起的工具调用会流式传回浏览器，在当前演示文稿上执行，再将结果返回模型。密钥始终留在服务器。

```ts
import { openai } from '@ai-sdk/openai';
// A backend route (framework-agnostic helpers from `ooxml-ui/pptx/ai`)
import { streamText } from 'ai';
import { buildPptxAiTools, buildPptxAiSystemPrompt } from 'ooxml-ui/pptx/ai';

export async function POST(request: Request) {
	const { messages } = await request.json();
	const result = streamText({
		model: openai('gpt-4o'),
		system: buildPptxAiSystemPrompt(),
		tools: await buildPptxAiTools(),
		messages,
	});
	return result.toUIMessageStreamResponse();
}
```

### `model`（浏览器端、本地模型或自带密钥） {#model-browser-side-local-bring-your-own-key}

将浏览器中创建的模型实例传给组件。助手直接在当前进程中运行工具循环，无需后端，适合本地模型、开发环境或可信环境。此方式的密钥存在于浏览器中，不适用于不可信的客户端。

```ts
import { createOpenAICompatible } from '@ai-sdk/openai-compatible';

const provider = createOpenAICompatible({
	name: 'local',
	baseURL: 'http://localhost:11434/v1', // e.g. Ollama
	apiKey: 'ignored-for-local',
});

const ai: PptxAiConfig = {
	connection: { kind: 'model', model: provider.chatModel('llama3.1') },
};
```

### `transport`（高级用法） {#transport-advanced}

提供已经构造好的 AI SDK `ChatTransport`，适用于 WebSocket 网关、工作流传输或测试。

```ts
const ai: PptxAiConfig = { connection: { kind: 'transport', transport: myTransport } };
```

## 审阅修改：提出建议与接受建议 {#reviewing-changes-the-propose-accept-flow}

默认的 `writePolicy: 'stage'` 会先**暂存**助手提出的编辑。面板显示建议卡片，包含简短的修改说明以及接受、拒绝和全部接受按钮。接受后，该修改作为单条记录进入组件的编辑历史，可以通过一次 Ctrl+Z 撤销；拒绝则丢弃建议。读取幻灯片、查找文本等只读工具和导航操作会立即执行，不经过暂存。

`writePolicy` 支持以下选项：

| 值                | 行为                                                                                     |
| ----------------- | ---------------------------------------------------------------------------------------- |
| `'stage'`（默认） | 写入先暂存，由用户明确接受或拒绝。                                                       |
| `'approve'`       | 使用 AI SDK 原生的逐次调用审批提示。                                                     |
| `'auto'`          | 立即应用写入，仍然支持撤销。`delete_slides` 等破坏性工具无论采用何种策略，始终需要确认。 |

## 助手可用的工具 {#what-the-assistant-can-do-tools}

助手默认使用一组固定的演示文稿工具。可以通过 `tools.enabled` 白名单或 `tools.disabled` 黑名单缩小范围，也可以通过 `tools.extra` 添加自定义工具。

这些工具共 **55 个**，由 `PptxAiToolName` 类型定义。大部分与 [MCP 服务器](/zh/packages/mcp)使用相同的函数，因此应用内助手与外部智能体使用相同工具名时语义一致。少数工具依赖当前 UI，例如选择、导航和解析后的演示文稿概览，因此只在组件中提供。

- **组件专用读取工具**（6 个）：`get_deck_overview`、`get_slide`、`get_element`、`get_speaker_notes`、`find_text`、`get_theme`
- **组件专用导航和辅助工具**（3 个）：`go_to_slide`、`select_elements`、`merge_tables`
- **其他读取工具**（7 个）：`get_metadata`、`get_layouts`、`find_placeholders`、`get_presentation_properties`、`run_accessibility_check`、`convert_to_markdown`、`export_to_json`
- **元素编辑工具**（13 个）：`add_element`、`update_element`、`delete_elements`、`arrange_elements`、`clone_element`、`set_element_animation`、`group_elements`、`ungroup_elements`、`batch_update_elements`、`update_element_style`、`replace_geometry`、`set_element_lock`、`manage_hyperlinks`
- **文本、表格、图表和 SmartArt 工具**（11 个）：`replace_text`、`manage_comments`、`update_table_cells`、`manage_table_structure`、`create_chart`、`update_chart`、`add_chart_series`、`remove_chart_series`、`update_chart_series_data`、`manage_smart_art`、`apply_template`
- **幻灯片结构工具**（6 个）：`add_slide`、`duplicate_slide`、`delete_slides`、`reorder_slides`、`update_slide_properties`、`set_slide_transition`
- **主题工具**（3 个）：`apply_theme_preset`、`update_theme_colors`、`update_theme_fonts`
- **演示文稿级工具**（6 个）：`set_canvas_size`、`update_metadata`、`manage_sections`、`update_presentation_properties`、`import_from_json`、`apply_layout`

演讲者备注可以通过 `get_speaker_notes` 读取，没有单独的写入工具；助手通过 `update_slide_properties` 修改备注。

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Read-only assistant: no editing tools.
	tools: {
		enabled: ['get_deck_overview', 'get_slide', 'get_element', 'find_text', 'go_to_slide'],
	},
};
```

## 演示文稿上下文 {#deck-context}

除了模型通过只读工具自行获取的信息，`contextStrategy` 还控制每轮对话主动发送多少演示文稿内容。

| 值                  | 每轮发送的内容                         |
| ------------------- | -------------------------------------- |
| `'outline'`（默认） | 整份演示文稿的精简逐页大纲。           |
| `'current-slide'`   | 仅当前幻灯片，以 Markdown 表示。       |
| `'none'`            | 不主动发送内容，模型完全依赖只读工具。 |

## 其他选项 {#other-options}

```ts
const ai: PptxAiConfig = {
	connection: { kind: 'endpoint', api: '/api/ai/chat' },
	// Appended to the assistant's base instructions (brand voice, guardrails).
	systemPromptExtras: 'Prefer concise edits. Never change the company logo.',
	// Persist and restore a conversation (host-owned storage).
	history: {
		load: (id) => myStore.load(id),
		save: (id, messages) => myStore.save(id, messages),
	},
	onError: (error) => reportToSentry(error),
};
```

## 隐私 {#privacy}

助手运行时，读取的幻灯片文本、结构以及 `contextStrategy` 允许的上下文，会发送给你配置的模型提供商。请根据数据策略选择连接方式和提供商。`endpoint` 模式让密钥和提供商选择保留在自己的基础设施中；`tools.enabled` 和 `contextStrategy: 'none'` 则可以限制助手读取的范围。

## 包体积 {#bundle-size}

AI 代码不会进入关键加载路径。只有传入 `ai` 配置，并且用户首次打开聊天面板时，才会动态导入面板和 `ai` SDK。未配置 AI 的预览器不会下载这些代码。

## `/ai` 入口 {#ai-entry}

每个绑定的 `/ai` 入口都导出 `createAiAssistant(config)`、句柄类型 `PptxAiAssistant` 以及 `PptxAi*` 配置类型。句柄携带该绑定自己的惰性面板加载器，也就是主入口处理普通配置时使用的同一个加载器。Angular 的句柄只携带配置：无论传入句柄还是普通配置，都由聊天服务加载运行时。

| 绑定    | 入口                     | 句柄携带的内容                                    |
| ------- | ------------------------ | ------------------------------------------------- |
| React   | `pptx-react-viewer/ai`   | `React.lazy` 聊天面板（其中加载 `@ai-sdk/react`） |
| Vue     | `pptx-vue-viewer/ai`     | `defineAsyncComponent` 聊天面板（`@ai-sdk/vue`）  |
| Svelte  | `pptx-svelte-viewer/ai`  | 聊天面板组件的加载器（`@ai-sdk/svelte`）          |
| Angular | `pptx-angular-viewer/ai` | 无（`AiChatService` 自行导入 `ooxml-ui/pptx/ai`） |
| Vanilla | `pptx-vanilla-viewer/ai` | 面板模块的加载器（`ooxml-ui/pptx/ai`）            |

### 本版本保证什么 {#what-this-release-guarantees}

在全部五个绑定中：

- 句柄是 `ai` 接受的、受支持的不透明值。传入句柄等同于传入它的配置：同一个面板，同样的行为。句柄不提供任何宿主扩展点：其类型没有任何成员，只有 `createAiAssistant` 能创建句柄，句柄也无法携带面板或你自己的任何代码。每个绑定都把句柄保存在包内私有的注册表中，任何入口都不导出它，因此绑定或 `ooxml-ui` 的其他任何导出都无法创建查看器会读取的句柄。唯一的例外属于内部实现：Angular 的 `/ai` 入口由 ng-packagr 单独构建，它通过 `createAiAssistantHandle` 访问该注册表。包根入口和 `/internals` 以 `@internal` 形式导出该函数，它只接受配置，创建的句柄与 `createAiAssistant` 创建的相同。
- `/ai` 入口是稳定的：`createAiAssistant`、`PptxAiAssistant` 和 `PptxAi*` 配置类型。
- 普通的 `PptxAiConfig` 继续受支持，不会被弃用。预览器会把任何不是句柄的值（包括形状与句柄相同的对象）当作普通配置，并像以前一样不经校验地交给面板；面板无法使用的配置会让面板进入错误状态。
- 由另一个绑定或另一份包副本（第二次安装，或混用 ESM 与 CommonJS 构建）创建的句柄也按同样方式处理。在开发环境中，预览器会对此警告一次（`process.env.NODE_ENV !== 'production'`；Angular 使用 `isDevMode()`）。
- 主入口仍然会到达惰性面板加载器。`pptx-react-viewer`、`pptx-vue-viewer`、`pptx-svelte-viewer` 及其 `/viewer` 入口，`pptx-angular-viewer`（以及同一文件的 `/internals`）和 `pptx-vanilla-viewer`，都通过与 `/ai` 相同的动态导入到达该绑定的 SDK 包和 `ooxml-ui/pptx/ai`。它们的初始代码块都不包含 `ai` SDK、`@ai-sdk/*`、`zod` 或自动化工具目录。因此句柄不会改变加载的内容和加载时机。

### 为什么现在提供句柄 {#why-the-handle-ships-now}

- 你可以在一个次版本中改用 `createAiAssistant`，这样下一个主版本就不需要你再修改代码。
- 句柄把助手放在一个由你主动选用的导入之后。在 React、Vue、Svelte 和原生 JavaScript 中，句柄携带的面板加载器来自 `/ai` 入口，而不是主入口。在 Angular 中，由包根入口中的聊天服务加载运行时，这样 `/ai` 就不会向包根入口传入任何代码。

### 下一个主版本会改变什么 {#what-the-next-major-release-changes}

在每个绑定中满足以下全部条件时，主入口才不再包含助手运行时：

- 主入口（`pptx-react-viewer`、`pptx-vue-viewer`、`pptx-svelte-viewer` 及其 `/viewer` 入口，`pptx-angular-viewer` 及其 `/internals`，`pptx-vanilla-viewer`）的源码与构建产物模块图，无论静态导入还是动态导入，都不会到达 `ai`、`@ai-sdk/*`、`zod`、`ooxml-ui/pptx/ai` 或 `ooxml-core/pptx/automation` 模块。
- 它们的 `.d.ts` 文件不会引用 `ai` 或 `@ai-sdk/*` 的任何类型。
- 打包绑定检查（`check-binding-packages.mjs`）要求每个绑定的入口中只有 `./ai` 会到达该绑定的 SDK 包（`@ai-sdk/react`、`@ai-sdk/vue`、`@ai-sdk/svelte`）和 `ooxml-ui/pptx/ai`。
- 主入口 `ai` 的类型中移除普通配置；运行时传入普通配置会抛出错误，并提示改用 `createAiAssistant`。

只有那个版本会改变行为：`ai` 只接受句柄，`PptxAi*` 配置类型（其中引用了 `ai` SDK 的类型）只从 `/ai` 入口导出，主入口也不再包含自己的惰性面板加载器。已经传入句柄的代码无需任何修改即可继续使用。在 Angular 中，运行时加载器目前位于包根入口，因此那个版本还必须把它移到 `/ai` 中，并且不能为此新增接受加载器的导出。
