# 框架绑定

六个轻量的原生集成都使用同一个浏览器元素和同一份 `mountViewer` 实现。每个集成都作为独立的、自包含的软件包发布；它们都不解析文档、不绘制图表，也不拥有其他渲染器。

每个原生句柄都提供 `createBlankDrawing(options?: CreateVsdxOptions)`。共享控制器会通过其现有的加载流程创建并接受新的源；状态钩子和文档事件照常更新。指针移动和框选在所有绑定中都运行于同一个共享 Web 组件内。

| 框架     | 软件包                 | 导出                                         |
| -------- | ---------------------- | -------------------------------------------- |
| Vanilla  | `visio-vanilla-viewer` | `mountViewer`                                |
| React    | `visio-react-viewer`   | `VisioViewer`                                |
| Vue 3    | `visio-vue-viewer`     | `VisioViewer`                                |
| Angular  | `visio-angular-viewer` | `VisioViewerComponent` (`visio-viewer-host`) |
| Svelte 5 | `visio-svelte-viewer`  | 默认导出 `VisioViewer`                       |
| Solid    | `visio-solid-viewer`   | `VisioViewer`                                |

只导入所需的框架软件包，以避免加载无关的框架运行时。这些适配器的源代码位于仓库的 `packages/bindings` 中，其 README 列出了用于保障它们的各项检查。

各绑定的连接方式是共享的。约定（`VIEWER_CONTRACT`、属性和事件键，以及句柄成员）以及 `mountFrameworkViewer` 的修订感知补丁转发，都位于 `packages/bindings/src/common.ts` 中。通用的绑定机制（延迟句柄、属性筛选，以及 React 和 Vue 的组件工厂）则是 `ooxml-ui/bindings` 工具包。React 和 Vue 基于这些工厂构建；Angular 从 `VIEWER_CONTRACT` 安装其句柄成员；Svelte 是一个组件；Solid 是一个 JSX 组件，直接挂载共享元素。`packages/bindings/tests/contract-parity.test.ts` 会依据约定检查每个适配器：属性相同，每个事件都经由回调映射和原生输出传递，每个句柄成员也都一致。

## 共享约定 {#shared-contract}

所有适配器都接受相同的属性：

- `document`：`VisioDocument` 或用于清除的 `null`
- `pageIndex`：从零开始的页面索引
- `zoom`：缩放比例，其中 `1` 表示 100%
- `showToolbar`：工具栏的可见性
- `canEdit`：编辑权限。完整查看器默认可编辑；`false` 使其只读（选择、搜索、缩放和导出仍然可用）。构建模块则默认为只读，参见[构建模块](/zh/building-blocks)
- `readOnly`：独立的限制。`true` 使图纸只读，并优先于 `canEdit`；`false` 或省略时，则只解除该限制
- `ribbonAddIns`：宿主在“帮助”之后添加的选项卡（参见[功能区加载项选项卡](/zh/api#ribbon-add-in-tabs)）
- `collaboration`：设置期间在房间中共享已打开的图纸，设为 `null` 则停止共享（参见[协作](/zh/collaboration)）
- `storage` 和 `customization`：宿主自有的保存和打开，以及共享的自定义对象（参见[自定义](/zh/api#customization)和[宿主自有存储](/api#host-owned-storage-and-file-commands)）
- `fileProviders`：“文件 > 打开”中的云端位置，例如 OneDrive（参见 [OneDrive](/zh/onedrive)）
- `events`：使用字面事件名称的类型化回调映射

事件映射目前包含 `document-load`、`document-change`、`document-error`、`page-change`、`zoom-change`、`shape-select`、`selection-change`、`office-ribbon-add-in`、`collaboration-change` 和 `file-command`。它派生自规范的约定。所有适配器都会传递每个事件；它们都不解析文档，也不拥有其他渲染器。

属性更新是局部的。原生框架的快照会进行差异比较，因此无关的重新渲染不会撤销命令式加载或用户更改的视口。省略或为 `undefined` 的属性会保留当前的查看器状态。`null` 则会明确清除文档。移除 `events` 属性会清除回调映射中的处理程序。若要进行外部替换，请提供新的文档对象；原地修改不是受支持的响应式机制。页面和缩放的交互可能会更新内部的查看器状态；当框架更新这些属性时，属性会请求相应的更改。

`document-change` 携带 `{ document, dirty, kind }`，其中 `kind` 为 `edit`、`undo`、`redo`、`remote`（协作者的更改）或 `save`（在 `markClean()` 之后）。命令式的 `ViewerHandle` 暴露 `element`、`controller`、`load(bytesOrBlob)` 和 `fit()`。编辑方法包括 `applyEdits(edits)`、`replacePlainText(pageId, shapeId, text)`、`duplicateSelection()`、`copySelection()`、`cutSelection()`、`pasteSelection()`、`undo()`、`redo()`、`cancelEdit()` 和 `exportVsdx()`。宿主保存导出的字节后，调用 `controller.markClean()`。协作通过属性驱动；`element.collaborationState`、`element.reconnectCollaboration()` 和 `element.stopCollaboration()` 位于句柄的 `element` 上。剪贴板准备状态可通过每个原生响应式订阅在 `controller.state.clipboard` 中获取；剪贴板方法需要浏览器的用户手势，并支持有界的查看器图形格式。它们会转发到同一套基于源、运行于 Worker 隔离环境中的实现。`ViewerEditState` 和 `VsdxExportResult` 是共享的导出类型；编辑状态可通过 `controller.state.edit` 获取。Angular 提供原生的 `documentChange` 输出；每个适配器都支持 `document-change` 回调映射。通过框架句柄进行的读取和操作，在挂载之前或卸载之后都会明确失败。所属组件卸载时，仍被保留的控制器或元素也会被释放。Vanilla 返回包含 `update()` 的完整共享绑定，以及可重复调用的 `destroy()`。

编辑器本身在客户端挂载。单独的 Node 测试会验证全部六个入口在没有浏览器全局对象的情况下可以导入；React、Vue、Svelte 和 Solid 还会在服务端渲染空的原生宿主。React 和 Vue 有测试会在挂载共享元素之前，对其实际的服务端标记进行水合。Angular 的服务端渲染、Solid 和 Svelte 的水合，以及跨文档挂载，仍未经验证。这些检查基于已安装的版本，并不意味着完整的 SSR 框架集成覆盖。

## 原生用法 {#native-usage}

React 使用转发的 ref 和普通的属性：

```tsx
const ref = useRef<ViewerHandle>(null);
<VisioViewer ref={ref} document={diagram} events={{ 'document-error': reportError }} />;
```

Vue 通过组件 ref 暴露相同的句柄，并同时发出原生的 kebab-case 事件：

```vue
<VisioViewer ref="viewer" :document="diagram" @document-error="reportError" />
```

Angular 暴露原生的输入和输出。将独立组件添加到应用的 imports 中：

```html
<visio-viewer-host [document]="diagram" (documentError)="reportError($event)" />
```

Angular 的输出包括 `documentLoad`、`documentChange`、`documentError`、`pageChange`、`zoomChange`、`shapeSelect`、`selectionChange`、`ribbonAddIn`、`collaborationChange` 和 `fileCommand`。`fileCommand` 只报告该命令；若要取消，Angular 宿主需从 `events['file-command']` 回调中返回 `false`。完整的 `events` 回调映射同样可作为输入接受。Vue 和 Angular 会先调用回调映射中的处理程序，然后再发出其原生输出；除非两种通知都需要，否则请只使用其中一种 API。

对于大型不可变的文档模型，请在 Vue 中使用 `shallowRef`/`markRaw`，在 Svelte 中使用 `$state.raw`，而不要在每次更新时对新构造的选项对象进行深层代理。将文档重新包裹在新的响应式代理中会改变其标识，并请求外部替换。不支持原地修改。

Svelte 通过 `events` 使用原生回调属性。其组件导出 `load`、`fit` 和 `getHandle`，因此无需暴露私有 DOM 即可访问完整的类型化句柄：

```svelte
<VisioViewer bind:this={viewer} document={diagram}
  events={{ 'document-error': reportError }} />
```

Solid 在挂载时通过 ref 回调提供句柄，在卸载或替换 ref 时则为 `undefined`：

```tsx
<VisioViewer
	document={diagram()}
	viewerRef={setViewer}
	events={{ 'document-error': reportError }}
/>
```

Vanilla 只拥有它自己创建的子元素：

```ts
const viewer = mountViewer(host, { document: diagram, zoom: 1 });
viewer.update({ events: { 'document-error': reportError } });
// Later:
viewer.destroy();
```

样式钩子因框架而异：React 使用 `className`/`style`，Vue 使用继承的属性，Angular 使用宿主属性，Svelte 和 Solid 使用 `class`/`style`。应用必须为宿主提供可用的高度。共享的查看器样式和无障碍功能则位于自定义元素中。
