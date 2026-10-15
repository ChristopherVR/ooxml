# 协作 {#collaboration}

“文件 > 共享”会在同一浏览器的多个窗口之间启动实时会话。跨设备共享暂不可用。

::: warning 部分支持，仅限同一浏览器
此版本没有托管服务，也没有服务器传输。任何内容都不会离开你的浏览器，也不会上传任何文档。
:::

## 目前已有的功能 {#what-exists-today}

- 基于 `ooxml-core/collab`（Yjs）并通过 `BroadcastChannel` 构建的实时会话，因此同一浏览器的窗口和标签页可以共享同一份图纸。
- 共享整个 VSDX 包，协作运行时仅在开始共享时才加载。
- 远程更改以一个可撤销的历史步骤到达，协作者通过共享的 `ooxml-ui` 在线状态元素显示。
- 会话名称为 1 至 64 个字母、数字、连字符或下划线。加入的窗口若发现该会话房间中已有图纸，则采用该图纸；否则由它为房间提供初始内容。

## 实时演示：两个窗口，任意框架 {#live-demo-two-windows-any-framework}

共享会话只是一个名称。下面的两个窗格是由不同框架绑定构建的独立演示环境，它们在加载时都会加入会话 `docs-collab`：窗口 A 是原生演示环境并打开示例，窗口 B 是 React 演示环境并从房间接收图纸。在一个窗格中编辑形状文本，另一个窗格会随之更新；撤销会还原远程步骤。

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer window A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer window B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### 任意搭配两个框架 {#pair-any-two-frameworks}

每个演示环境都接受 `?share=<session>`，它会填入“文件 > 共享”并启动共享。只有会话名称起作用，因此可以任意搭配：主页的[实时演示](/zh/#live-demo)为每个窗口提供了框架选择器；在两个标签页中，你也可以打开例如 `/demo-vue/?sample=1&share=my-room` 和 `/demo-angular/?share=my-room`。你还可以在任意两个窗口中手动操作：文件，然后共享，再输入相同的会话名称。

### 为什么仅限同一浏览器 {#why-it-is-same-browser-only}

该会话使用 `ooxml-core/collab`（Yjs）并通过 `BroadcastChannel` 工作。`BroadcastChannel` 只连接同一来源、同一浏览器配置文件中的窗口、标签页和框架，无论它们由哪种框架构建，除此之外别无其他。两台不同的机器无法以这种方式互通；那需要由宿主应用程序提供服务器传输，而此版本未附带任何服务器传输。

## 支持与不支持的功能 {#what-is-and-is-not-supported}

| 支持                                           | 不支持                                         |
| ---------------------------------------------- | ---------------------------------------------- |
| 基于 Yjs、通过 BroadcastChannel 的同浏览器会话 | 跨设备共享：需要宿主提供服务器传输，目前未附带 |
| 整包共享，每次远程更改对应一个可撤销步骤       | 字段级合并：并发更改不会合并，以最后一个包为准 |
| 使用共享在线状态元素显示协作者                 | 批注、身份验证和角色                           |
| 以“文件 > 共享”作为入口                        | 面向自定义传输的已文档化嵌入 API               |

相关证据记录在[功能清单](/parity)的“Collaboration and comments”一节中。
