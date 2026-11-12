# OneDrive 集成

如需分步设置（Microsoft Entra 注册、登录和完整示例），请先阅读
[OneDrive 快速入门](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-quickstart.md)。

## 文件 > 打开 > OneDrive {#file-open-onedrive}

将提供方作为 `fileProviders` 传入，“文件 > 打开”就会在“此设备”旁边显示一个 OneDrive 位置：

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken, // your sign-in owns the token
		signIn,
		canWrite: async (file) => file.driveId === (await myDriveId()),
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

该位置会列出最近的文件和“我的文件”，支持搜索，并可打开共享链接。所选工作簿的加载方式与其他工作簿相同，通过“保存”会保存回同一项目（“另存为”则保存为副本）；只有在上传成功完成后才会被标记为已保存，上传失败则保持未保存状态。`canWrite` 拒绝的文件（其他驱动器或共享文件）会以只读方式打开。`fileProviders` 是每个框架绑定的属性（React、Vue、Solid、Svelte、Angular 和原生 JavaScript 中均为 `fileProviders`）。演示应用（`demos/xlsx`）默认传入演示用的 OneDrive 提供方；可通过 `demos/xlsx/demo-vanilla/.env.local` 中的 `VITE_ONEDRIVE_CLIENT_ID` 进行配置，或在其登录对话框中粘贴 Graph Explorer 令牌。下面较底层的 `storage` 配置适用于自行构建打开界面的宿主。

同一存储契约也支持自定义提供方。有关 `createStorageSession`、按格式的合并覆盖以及套件引导设置，请参见
[与提供方无关的存储和合并配置](https://github.com/ChristopherVR/ooxml/blob/main/docs/storage-integrations.md)。
所有 Excel 框架软件包都可以使用这些扩展点。

Excel 使用共享的 `ooxml-core/embedding/onedrive` 连接。请在宿主中解析共享链接，验证项目特定的权限，并将下载的 XLSX 字节交给现有编辑器：

```ts
// opened and session come from createOneDriveClient / createOneDriveSession.
editor.storage = session.storage;
editor.readOnly = opened.file.access !== 'write';
await editor.load(opened.bytes, opened.file.name);
```

元素和框架句柄提供 `saveBytes('xlsx')`、`load(bytes, fileName?)` 和 `markClean()`。宿主自定义保存时调用 `session.save(await editor.saveBytes('xlsx'))`；只有在持久化成功且之后没有更新的编辑时，才能标记为已保存。请使用常规的编辑器保存路径来提交活动单元格的编辑。CSV 是单独的导出格式，绝不能替换已连接的 XLSX 项目。
除非宿主阻止了可取消的 `file-command` 事件，否则原生的“保存”会委托给 `storage`。

同一套 `storage` 和 `readOnly` 契约同样适用于 `@christophervr/xlsx-react-viewer`、
`xlsx-vue-viewer`、`xlsx-angular-viewer`、`xlsx-svelte-viewer`、`xlsx-solid-viewer` 和
`xlsx-vanilla-viewer`，以及普通的 `ooxml-ui/xlsx` 元素。`@christophervr/xlsx-core` 是模型门面，而非身份验证服务。

有关身份验证、共享以及同时保留本地和远程版本的恢复方法，请参见
[Office OneDrive 软件包参考](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#excel-xlsx)。
[测试指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)说明了无需凭据的测试夹具和真实账户验收测试；本项目不提供任何模拟界面或服务。

将 `ooxml-core/xlsx/collab` 中的 `mergeWorkbookSnapshots` 注册为会话的 `merge` 回调，即可在保存时合并受支持的独立工作簿编辑。套件会自动启用它；但重叠的单元格编辑和不受支持的更改仍需要恢复。这并不会将工作簿协作会话连接到 Microsoft Excel 的共同创作，也不会刷新外部数据，或更改[功能限制](/zh/features)。

## 账户 > 已连接的服务 {#account-connected-services}

同一个 `fileProviders` 还会在文件视图的**账户**页面中填充“已连接的服务”列表（共享的 `<office-ui-connected-services>`）：每个提供方对应一行，显示已登录的账户以及“连接”或“断开连接”按钮，分别调用提供方自身的 `signIn` 和 `signOut`。未提供 `signOut` 的提供方会显示由宿主管理其会话。该列表会跟随提供方的 `subscribe`，因此在“文件 > 打开”中登录也会更新它，其文字来自各语言的 Excel 词典。没有 `fileProviders` 时则不会有账户页面。

## 从自己的界面打开云端文件 {#open-a-cloud-file-from-your-own-ui}

只有打开文档之后，才能访问“文件 > 打开”。若要从起始页或自己的选择器打开文件，请从任意 `CloudFileProvider` 获取 `CloudOpenedFile`（共享的 `<office-ui-cloud-files>` 浏览器会在 `office-cloud-file-open` 中发出它），然后在 `<xlsx-editor>` 上将其传给 `editor.openCloudFile(file)`（每个绑定的句柄也都提供此方法）。它的运行路径与“文件 > 打开 > OneDrive”相同：使用文件名；当 `canEdit` 为 false 时以只读方式打开；在另一份文档加载之前，“保存”会路由到该文件自身的存储。
