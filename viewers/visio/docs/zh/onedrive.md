# OneDrive 集成

如需按步骤完成设置（Microsoft Entra 注册、登录和完整示例），请先阅读 [OneDrive 快速入门](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md)。

如需对接其他服务或自定义的 OneDrive 实现，请参阅[与提供方无关的存储集成](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)。`createStorageSession` 会连接到所有 Visio 软件包中的同一个 `storage` 属性，并提供宿主自有的修订检查、合并处理程序、冲突界面和共享策略。

## 通过云端位置打开文件 {#file-open-with-cloud-places}

传入 `fileProviders`（元素的属性，也是每个绑定的 prop），即可在“文件 > 打开”中添加云端位置。当该列表不为空时，“打开”会显示**此设备**（通常的浏览按钮）以及每个提供方对应的一个位置。每个位置都会渲染共享的 `<office-ui-cloud-files>` 浏览器，用于 `.vsdx` 图纸：登录、最近使用、我的文件、搜索，以及通过共享链接打开。所有语言中的文字都取自 Visio 的显示语言。

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeDrawingSnapshots } from 'ooxml-core/visio/collab';

const oneDrive = createOneDriveFileProvider({
	getAccessToken, // the host's own sign-in
	signIn, // opens the host's sign-in UI and resolves with the account
	canWrite: async (file) => file.driveId === (await myDriveId()),
	mergers: {
		vsdx: ({ base, local, remote }) => mergeDrawingSnapshots(base.bytes, local, remote.bytes),
	},
});
// <VisioViewer fileProviders={[oneDrive]} /> (React); [fileProviders] in Angular; :file-providers in Vue.
```

在那里选择的图纸会通过常规的 `.vsdx` 路径加载，并采用该文件的名称，然后：

- 若文件本身为只读（`canEdit` 为 false），则图纸也为只读。若宿主已将查看器设为只读（`canEdit`、`readOnly`、协作的 `viewer` 角色或 `capabilities.edit`），则对于可编辑的文件，它仍保持只读：两者中较严格者生效。
- “文件 > 保存”和“另存为”会使用该文件自己的存储，而不是宿主的 `storage`，直到加载另一份图纸（打开、新建、`load` 或 `document`）为止。只有在提供方保存了字节之后，保存才会完成：此时图纸显示为已保存；若上传失败，图纸则保持未保存状态，并触发 `document-error`。
- 可取消的 `file-command` 事件仍会先行触发（`open`、`save`）；取消 `open` 会使当前图纸保持不变。图纸打开后，后台视图随即关闭。
- 导出诊断与任何其他保存一样进行报告。旧版 `.vsd` 不是云端编辑的目标。

演示默认提供一个 OneDrive 位置，通过仅用于演示的对话框登录（输入客户端 ID，或将 Graph Explorer 令牌粘贴到**访问令牌**中）。该对话框属于演示配置，而非产品界面；请参阅 `demos/onedrive`。若要跳过该对话框，请在 `demos/visio/.env.local` 中设置 `VITE_ONEDRIVE_CLIENT_ID`（以及可选的 `VITE_ONEDRIVE_AUTHORITY`），在 `viewers/visio` 中运行 `bun run dev`，并将其打印出的地址（末尾带斜杠）注册为单页应用的重定向 URI。

## 账户 > 已连接的服务 {#account-connected-services}

同一个 `fileProviders` 还会填充文件视图中的**账户**页面，显示“已连接的服务”列表（共享的 `<office-ui-connected-services>`）：每个提供方一行，显示已登录的账户以及“连接”或“断开连接”按钮，分别调用提供方自己的 `signIn` 和 `signOut`。若提供方没有 `signOut`，则会显示会话由宿主管理。该列表会跟随提供方的 `subscribe`，因此在“文件 > 打开”中登录也会更新此列表，其文字取自所有语言的 Visio 词典。该页面也保留本地配置文件和产品信息。与 Word 和 Excel 一样，只有存在 `fileProviders` 时才会显示账户页面；没有它时，配置文件仍可在**文件 > 选项**中编辑。

## 直接使用会话 {#using-a-session-directly}

Visio 对原生 VSDX 文件使用共享的 `ooxml-core/embedding/onedrive` 连接。Microsoft 登录、链接解析和按项目验证的权限均由宿主负责。将下载的字节交给现有的查看器，并在访问权限未知时加上宿主的只读限制：

```ts
// Set these through the element or binding's existing props/options.
const canEdit = opened.file.access === 'write';
const options = { storage: session.storage, canEdit, readOnly: !canEdit };
// After mounting with options:
await viewer.load(opened.bytes);
```

`exportVsdx()` 返回一个包含原生 `bytes` 和写入器诊断信息的对象。自定义的“保存”按钮可以调用 `await session.save((await viewer.exportVsdx()).bytes)`。除非宿主阻止了可取消的 `file-command` 事件，否则原生的“文件 > 保存”使用相同的存储约定。请查看诊断信息，而不要承诺无损导出。二进制 `.vsd` 预览不是原生的 OneDrive 编辑目标。

这适用于 `ooxml-ui/visio` 元素，以及 `visio-react-viewer`、`visio-vue-viewer`、`visio-angular-viewer`、`visio-svelte-viewer`、`visio-solid-viewer` 和 `visio-vanilla-viewer`。`visio-core` 是模型门面，该连接不会为其添加 Microsoft 身份验证。

软件包的连接方式、身份验证、考虑权限的共享以及保留两个版本的冲突恢复，请参阅 [Office OneDrive 参考](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#visio-vsdx)。测试本地传输和实时服务验收，请参阅[测试](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)。本项目不附带任何模拟界面或服务。

将 `ooxml-core/visio/collab` 中的 `mergeDrawingSnapshots` 注册为会话的 `merge` 回调，即可在保存时合并受支持的、相互独立的形状和属性编辑。套件会自动这样做。重叠或含糊不清的更改仍需要恢复。这并非 Microsoft 的多人协作或文件锁定。[Visio 的限制](/parity)和导出诊断同样适用。

## 从你自己的界面打开云端文件 {#open-a-cloud-file-from-your-own-ui}

只有在已经打开文档之后，才能使用“文件 > 打开”。若要从起始页或你自己的选择器中打开文件，请从任一 `CloudFileProvider` 获取 `CloudOpenedFile`（共享的 `<office-ui-cloud-files>` 浏览器会在 `office-cloud-file-open` 中发出该对象），然后将其传给 `<visio-viewer>` 上的 `viewer.openCloudFile(file)`（每个绑定的句柄也都提供此方法）。其执行路径与“文件 > 打开 > OneDrive”相同：采用文件名称，`canEdit` 为 false 时只读，保存则路由到该文件自己的存储，直到加载另一份文档为止。
