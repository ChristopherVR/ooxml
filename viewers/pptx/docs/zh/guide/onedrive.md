---
title: OneDrive 集成
description: 将每个 PowerPoint 查看器绑定连接到 Office 共享的 OneDrive 存储、共享和冲突恢复功能。
---

# OneDrive 集成

每个 PowerPoint 框架包都接受通过现有的 `storage` 属性传入的自定义存储提供程序。有关 `createStorageSession`、身份验证和传输层覆盖、按格式的合并处理程序、宿主冲突对话框以及显式共享选择，请参阅[与提供程序无关的集成与配置](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)。

PowerPoint 与 Word、Excel 和 Visio 使用相同的可选连接 `ooxml-core/embedding/onedrive`。Microsoft 身份验证和已验证的项目权限由宿主提供。这里不包含内置的 Microsoft 应用注册、模拟服务或模拟的 OneDrive 演示。

如需分步设置指南（包括应用注册、登录、共享链接框和一个完整的 React 示例），请先阅读 [OneDrive 快速入门](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md)。每个框架演示都在文件 > 打开下提供 OneDrive，供你用自己的 OneDrive 试用；快速入门中的“在 PowerPoint 演示中试用”一节说明了设置方法。

## 文件 > 打开中的位置：`fileProviders` {#file-open-places-fileproviders}

传入 `fileProviders`（一个 `CloudFileProvider` 列表，例如来自 `ooxml-core/embedding/onedrive` 中的 `createOneDriveFileProvider`），打开页面就会列出**此设备**以及每个提供程序对应的一个位置。该属性在所有绑定中拼写一致：`fileProviders`（React、Angular 输入属性、Svelte、vanilla 选项）以及 `:file-providers`（Vue）。

```tsx
<PowerPointViewer content={bytes} canEdit fileProviders={[oneDriveProvider]} />
```

若要合并其他人在此期间保存的编辑，请为提供程序提供 PowerPoint 合并器。没有合并器时，较新的 OneDrive 版本总会打开冲突选项；`merge: false` 对所有格式都会以同样的方式询问；`maxMergeAttempts`（适用合并时默认为 3）限制重试次数：

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergePptxSnapshots } from 'ooxml-core/pptx';

const oneDriveProvider = createOneDriveFileProvider({
	getAccessToken,
	canWrite,
	signIn,
	mergers: {
		pptx: ({ base, local, remote }) => mergePptxSnapshots(base.bytes, local, remote.bytes),
	},
});
```

从某个位置打开的文件会像其他打开操作一样加载（撤销历史重置，并触发 `onLoad`）。此后，查看器会使用该文件自己的存储执行文件 > 保存、另存为和恢复副本，而不再使用 `storage` 属性；如果服务不允许编辑，则以只读方式打开。`canEdit={false}` 或 `readOnly` 的宿主仍保持只读。该状态会一直保持，直到加载其他文档为止：新建演示文稿、从此设备或“最近使用”打开文件，或宿主提供新的 `content`。保存失败时会显示保存失败通知，状态栏保留**未保存的更改**（Unsaved changes）。

## 账户 > 已连接的服务 {#account-connected-services}

同一个 `fileProviders` 还会在**文件 > 账户**页面的个人资料、存储和登录卡片下方填充**已连接的服务**（Connected services）部分。每个提供程序对应一行，显示其账户（名称和电子邮件）以及**连接**（Connect）或**断开连接**（Disconnect）按钮，分别调用提供程序自己的 `signIn` 和 `signOut`。没有 `signOut` 的提供程序会显示其会话由贵组织管理。该列表跟随提供程序的 `subscribe`，因此在文件 > 打开中登录会同步更新列表。没有 `fileProviders` 时不会出现额外内容，现有的 `accountAuth` 登录卡片保持不变。五个绑定都会渲染共享的 `office-ui-connected-services` 元素（显示为 `pptx-ui-connected-services`），并带有翻译后的标签。

[完整的 Office 集成指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)记录了身份验证、会话存储、安全的云端副本、共享角色和范围、冲突恢复以及每种包类型的参考说明。

```ts
import { createOneDriveClient, createOneDriveSession } from 'ooxml-core/embedding/onedrive';

const drive = createOneDriveClient({
	getAccessToken: () => microsoftSession.getGraphAccessToken(),
	canWrite: (file) => permissionService.canUpdate(file.driveId, file.itemId),
});
const opened = await drive.open(sharingLink);
const session = createOneDriveSession(drive, opened);
// Pass session.storage and opened.file.access === 'write' to the viewer.
```

本示例中的会话和权限服务属于宿主集成，并非内置对象。请将适配器作为 `ooxml-core` 的直接依赖导入。`ooxml-ui/pptx/onedrive` 仅作为早期 PowerPoint 专用导入方式的兼容接口保留。

## 五个绑定 {#all-five-bindings}

| 绑定/包                        | 打开字节                              | 限制编辑                      | 导出当前 PPTX                     |
| ------------------------------ | ------------------------------------- | ----------------------------- | --------------------------------- |
| React，`pptx-react-viewer`     | `content={opened.bytes}`              | `canEdit={canEdit}`           | `await ref.current.getContent()`  |
| Vue，`pptx-vue-viewer`         | `:content="opened.bytes"`             | `:can-edit="canEdit"`         | `await viewer.value.getContent()` |
| Angular，`pptx-angular-viewer` | `[content]="opened.bytes"`            | `[canEdit]="canEdit"`         | `await viewer.getContent()`       |
| Svelte，`pptx-svelte-viewer`   | `source={opened.bytes}`               | `editable={canEdit}`          | `await viewer.getContent()`       |
| Vanilla，`pptx-vanilla-viewer` | `await viewer.loadFile(opened.bytes)` | `viewer.setEditable(canEdit)` | `await viewer.save()`             |

`canEdit`（Svelte 中为 `editable`）表示 `opened.file.access === 'write'`。访问权限未知时保持只读。每个绑定都以现有的属性、输入属性或选项形式接受 `storage`。将 `session.storage` 赋给它，即可把原生的文件保存连接到 OneDrive。`onOpenFile` 可以打开宿主的共享链接选择器；请勿把 Microsoft 共享页面的 URL 传给通用的文档源加载器。

只有在 `storage.save` 解析之后，查看器才会将演示文稿标记为已保存，而且只有在上传期间没有发生编辑时才会如此。如果运行中的保存失败，则演示文稿保持未保存状态。对于由宿主控制的“保存”按钮，请导出当前 PPTX 并调用 `session.save(bytes)`。保持编辑器挂载并保留撤销历史。上传期间的编辑应与已上传的快照分开跟踪。不要把原生保存方法的解析视为保存成功的证明：某些方法会通过 `onError` 和通知栈报告失败。参见[宿主自有存储](/zh/guide/host-owned-storage)。

## 冲突与共享 {#conflicts-and-sharing}

仅修订号发生变化时，如果下载的云端字节仍与已加载的基线相同，可以自动重试一次。真正的内容冲突会在 `OneDriveConflictError` 中保留基线、本地和远程版本。系统会提供**将我的编辑另存为 OneDrive 副本**（Save my edits as a OneDrive copy）、**保留本地副本并打开最新版本**（Keep local copy and open latest）、**下载我的编辑**（Download my edits）和**继续编辑**（Keep editing）选项。请勿静默覆盖或重新加载。将 `mergePptxSnapshots` 注册为会话的 `merge` 回调，即可自动合并受支持的独立幻灯片编辑。套件默认会这样做。不受支持的重叠编辑仍需要恢复操作；会话不会在最终上传提交之前锁定文件。

将文件另存为兼容的原生 PPTX 时，会通过会话存储创建单独的云端副本。导出为 PPT、PPTM、PPSX 或其他格式时，不得替换原始 PPTX，应作为单独的显式导出流程保留。只有在用户选择了查看/编辑角色和范围后，“共享”操作才会创建链接。共享 OneDrive 链接并不会将 PowerPoint 协作室连接到 Microsoft PowerPoint 联合编辑。

另请参阅[集成测试](/zh/guide/onedrive-testing)和现有的[功能限制](/zh/guide/limitations)。
