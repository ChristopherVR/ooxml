# 宿主自有存储

如果需要基于现有文件的 Microsoft Graph 提供程序，并附带无需应用注册的演示，请参阅 [OneDrive 集成](/zh/guide/onedrive)和 [OneDrive 测试](/zh/guide/onedrive-testing)。

每个绑定都接受一个 `storage` 属性（Angular 为输入属性，Vanilla 为选项），它把查看器的文件读写交给宿主处理。它是 `ooxml-core/embedding` 中的 `DocumentStorage` 合约，与 Word、Excel 和 Visio 编辑器接受的合约相同，并通过 `ooxml-ui/pptx` 以类型形式重新导出。

```ts
import type { DocumentStorage } from 'ooxml-ui/pptx';

const storage: DocumentStorage = {
	// Required: File > Save and Save As.
	async save(bytes, { fileName, mimeType, reason }) {
		await uploadToMyServer(fileName, bytes, mimeType);
		// Optional: report the name the document now has.
		return { fileName };
	},
	// Optional, all three together: AutoRecover copies.
	readRecovery: (id) => myRecovery.get(id),
	writeRecovery: (entry) => myRecovery.put(entry),
	clearRecovery: (id) => myRecovery.delete(id),
	// Optional: lets Version History list the kept copies.
	listRecovery: () => myRecovery.list(),
};
```

## 保存与另存为 {#save-and-save-as}

如果没有提供 `storage`，查看器会下载保存后的文件。提供了 `storage` 后，保存以及每种另存为格式（`.pptx`、`.ppsx`、`.pptm`、`.ppt`）都会改为调用 `storage.save(bytes, meta)`：

- `meta.fileName` 由源文件的主文件名加上所写格式的扩展名组成（`savedPresentationFileName`），因此 `report.ppt` 会保存为 `report.pptx`。
- `meta.mimeType` 与格式对应（`.ppt` 使用 OLE2 类型，`.ppsx` 和 `.pptm` 各有自己的类型）。
- `meta.reason` 表示所执行的命令：文件 > 保存（标题栏、后台视图中的“保存”项、移动端工具栏以及 Vanilla/Svelte 的 `downloadPptx()`）为 `'save'`，所有另存为条目（包括另存为 > PowerPoint 演示文稿 (.pptx)）为 `'saveAs'`。

保存失败（宿主的 `save()` 被拒绝，或序列化失败）在五个绑定中的处理方式相同：绑定的 `error` 事件（`onError`、`@error`、`(error)`、`onerror`）会收到 `Save failed: <reason>`，用户会在查看器的通知栈中看到“Save failed”提示，并且保留自动恢复副本（只有保存成功后才会丢弃该副本，选项 > 保存中要求丢弃时除外）。保存方法总是解析（resolve），不会拒绝（reject）。

写出的字节与 `getContent()` 返回的相同，因此受密码保护的演示文稿会以加密形式交给宿主。

## 恢复副本 {#recovery-copies}

恢复自动保存（以 `filePath` 为键的 `autosave` 和 `autosaveIntervalMs` 属性）通常写入库自己的 IndexedDB 数据库。当 `storage` 提供 `readRecovery`、`writeRecovery` 和 `clearRecovery` 时，所有恢复路径都会改用它们：自动保存写入器、加载时显示的恢复提示（及其“丢弃”按钮）、保存成功后的清理、版本历史、文件 > 打开 > 最近使用、账户 > 存储与隐私，以及 `restoreSessionDeck(storage)`。条目 id 为 `filePath` 键，`savedAt` 为毫秒时间戳。如果存储只提供了这三个方法中的部分，则恢复功能会忽略它，仍使用 IndexedDB。

列表功能（最近使用、账户用量和版本历史）需要 `listRecovery`，没有它这些列表为空。选项 > 保存中的清理功能（“删除缓存文件”、“关闭时清除缓存”以及“缓存保留”清扫）只有在宿主提供 `listRecovery` 时才会清除宿主的副本。没有它，查看器不会处理宿主的副本，保留策略由宿主负责。

## 没有 `file-command` 事件 {#no-file-command-event}

Word、Excel 和 Visio 元素会在保存、打开和新建之前触发可取消的 `file-command` 事件。PowerPoint 绑定没有此事件：若要接管保存，请传入 `storage`；若要接管文件 > 打开，请使用 `onOpenFile`（Vue 和 Angular 为 `onOpenFile` 输入属性）。
