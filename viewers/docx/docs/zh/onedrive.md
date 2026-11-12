# OneDrive 集成

如需逐步的设置说明（Microsoft Entra 应用注册、登录以及完整示例），请先阅读 [OneDrive 快速入门](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md)。

对于其他服务或自定义的 OneDrive 流程，请使用相同的 `storage` 属性，并参阅 [`createStorageSession` 与提供方无关的集成指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)。宿主可以替换身份验证、传输、合并处理程序、冲突界面和共享策略，而无需更改任何 Word 框架绑定。

所有 Word 包都使用共享的 `ooxml-core/embedding/onedrive` 连接。宿主负责解析共享链接、提供 Microsoft 身份验证并确认写入权限，然后将生成的会话连接到现有的编辑器存储契约。

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
editor.fileName = opened.file.name;
await editor.load(opened.bytes);
```

`ooxml-ui/docx` 中的普通 `<docx-editor>` 提供 `saveBytes()`。每个框架句柄都提供 `save()`（返回 `Blob`）、`load(bytes)`、`markClean()` 和 `dirty`。自定义的保存按钮可以将 Blob 转换为字节，并调用 `session.save(bytes)`；只有在持久化成功且期间没有发生更新的编辑时，才应标记为干净。除非宿主阻止了可取消的 `file-command` 事件，否则原生保存会自动通过 `storage` 进行。

## 在“文件 > 打开”中从 OneDrive 打开 {#open-from-onedrive-in-file-open}

将 `fileProviders`（元素的属性，也是每个绑定的属性）设置为一个 `CloudFileProvider` 列表后，“文件 > 打开”会显示多个位置：“此设备”（浏览按钮），以及每个提供方对应的一个位置。每个位置都会渲染共享的 `<office-ui-cloud-files>` 浏览器（最近使用、我的文件、搜索、通过共享链接打开），并过滤为 `.docx`。其标签来自每种显示语言中的 Word 词典。

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDocxSnapshots } from 'ooxml-core/docx';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your Microsoft authentication
		signIn, // opens your sign-in UI and resolves with the account
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			docx: ({ base, local, remote }) => mergeDocxSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

在某个位置中选择的文件会像其他文档一样加载，并采用该文件的名称；当提供方报告 `canEdit: false` 时，文件为只读（取宿主的 `readOnly` 与文件设置中限制更严格的一方）。“文件 > 保存”和“另存为”会写入该文件自身的存储，而不是宿主的 `storage`，直到加载另一个文档为止；标题栏只有在上传完成后才显示已保存。上传失败时，文件保持未保存状态，并触发 `document-error`。从某个位置打开文件不会触发 `file-command`（由宿主选择了提供方）；但“保存”和“另存为”仍会触发，宿主若取消这些事件，则由其自行处理。宿主若移除 `open` 能力，则会连同位置一起隐藏该页面。

Word 演示默认传入 `demos/onedrive` 中的 `createDemoOneDriveProvider`。要从“文件 > 打开 > OneDrive”登录：在 `demos/docx/demo-vanilla/.env.local` 中设置 `VITE_ONEDRIVE_CLIENT_ID`（以及可选的 `VITE_ONEDRIVE_AUTHORITY`），即可通过 MSAL 使用你自己的 Entra 应用注册；也可以在演示对话框中输入客户端 ID，或将 Graph Explorer 令牌粘贴到其“Access token”字段中。请将演示的源（带末尾斜杠）注册为单页应用的重定向 URI。演示只会编辑已登录用户自己云盘中的文件；其他文件以只读方式打开。

这适用于 `docx-react-viewer`、`docx-vue-viewer`、`docx-angular-viewer`、`docx-svelte-viewer`、`docx-solid-viewer` 和 `docx-vanilla-viewer`。它们各自转发相同的 `storage` 和 `readOnly` 属性/选项。`docx-core` 是模型门面，不实现 Microsoft 登录。

[Office OneDrive 参考文档](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx)包含包矩阵、身份验证设置、共享范围以及可操作的冲突恢复方法。[测试指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)区分了测试专用的传输与真实的 Microsoft 验收测试。本项目不提供模拟界面或模拟服务。

将 `ooxml-core/docx` 中的 `mergeDocxSnapshots` 注册为会话的 `merge` 回调，以便在保存时合并相互独立的文本和属性编辑。套件会自动启用此功能。结构不明确或替换内容不兼容时，会同时保留两个版本以便恢复。这并不会将 Word 协作房间连接到 Microsoft Office 共同创作。现有的[保留限制](/zh/preservation)仍然适用。

## 账户 > 已连接的服务 {#account-connected-services}

同一个 `fileProviders` 还会填充文件视图中的**账户**页面，显示“已连接的服务”列表（共享的 `<office-ui-connected-services>`）：每个提供方一行，显示已登录的账户以及“连接”或“断开连接”按钮，分别调用提供方自身的 `signIn` 和 `signOut`。没有 `signOut` 的提供方会显示由宿主管理其会话。该列表会跟随提供方的 `subscribe`，因此在“文件 > 打开”中登录也会更新它；其文字来自各语言的 Word 词典。没有 `fileProviders` 时，不会出现账户页面。

## 从你自己的界面打开云端文件 {#open-a-cloud-file-from-your-own-ui}

只有在文档打开之后，才能使用“文件 > 打开”。要从起始页或你自己的选择器中打开文件，请从任意 `CloudFileProvider` 获取一个 `CloudOpenedFile`（共享的 `<office-ui-cloud-files>` 浏览器会在 `office-cloud-file-open` 中发出它），然后将其传给 `<docx-editor>` 上的 `editor.openCloudFile(file)`（每个绑定的句柄也都暴露了该方法）。它与“文件 > 打开 > OneDrive”走相同的路径：使用该文件的名称；当 `canEdit` 为 false 时为只读；在加载另一个文档之前，保存会路由到该文件自身的存储。
