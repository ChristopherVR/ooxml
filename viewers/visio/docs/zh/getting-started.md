# 快速开始 {#getting-started}

在本地打开图表，了解当前的限制，并将同一个查看器嵌入到你的应用程序中。

::: warning 初期查看器支持，并非 Visio 认证
当前为公开测试版。请勿将文件成功打开视为每个形状、公式或格式功能都已被忠实呈现的证据。请参阅[功能清单](/parity)。
:::

## 试用本地演示环境 {#try-the-local-playground}

1. [打开演示环境](/demo/){target="_self"}，查看一份原创的示例文档。
2. 通过文件选择器或拖放区域选择本地 `.vsdx` 文件。
3. 切换页面、调整缩放、适应页面并选择形状。
4. 查看兼容性说明。请将重要文件与其在 Microsoft Visio 中的渲染效果进行对比。

不会上传任何文档。这些控件支持搜索、显示图层、实验性的纯文本编辑、有限的撤销/重做、SVG 导出以及明确的 VSDX 副本下载。打印快照仅作为 API 层面的准备工作，不提供打印对话框。通用绘图、PDF 导出以及大多数旧版二进制 `.vsd` 功能不在此版本范围内。

同一演示环境的其他框架版本列在[演示页面](/demos)中。

## 安装框架包 {#install-a-framework-package}

为你的框架安装一个查看器包。每个包都包含共享查看器、其适配器、Worker 和文档 API。

::: code-group

```bash [React]
npm install visio-react-viewer react
```

```bash [Vue]
npm install visio-vue-viewer vue
```

```bash [Angular]
npm install visio-angular-viewer
```

```bash [Svelte]
npm install visio-svelte-viewer svelte
```

```bash [Solid]
npm install visio-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install visio-vanilla-viewer
```

:::

| 框架     | 包                     | 导出                   |
| -------- | ---------------------- | ---------------------- |
| Vanilla  | `visio-vanilla-viewer` | `mountViewer`          |
| React    | `visio-react-viewer`   | `VisioViewer`          |
| Vue 3    | `visio-vue-viewer`     | `VisioViewer`          |
| Angular  | `visio-angular-viewer` | `VisioViewerComponent` |
| Svelte 5 | `visio-svelte-viewer`  | 默认导出 `VisioViewer` |
| Solid    | `visio-solid-viewer`   | `VisioViewer`          |

六个适配器都通过共享查看器转发相同的控制器、属性、事件和命令式句柄。请参阅[框架指南](/frameworks/react)和[绑定约定](/bindings)。

## 挂载查看器 {#mount-a-viewer}

容器必须有高度。可以将浏览器的 `File`、`Blob`、`Uint8Array` 或 `ArrayBuffer` 传给 `load`。

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	events: { 'shape-select': (shape) => console.log(shape) },
});

await viewer.load(file); // A local File or bytes
viewer.fit();

// When your view unmounts:
viewer.destroy();
```

## 运行开发工作区 {#run-the-development-workspace}

请使用 Node.js 22.12 或更高版本。常规构建会从 npm 安装已发布的 `ooxml-core/visio` API。

```bash
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run dev
```

使用 Vite 打印的本地 URL。这会启动开发服务器，不会发布站点。

## 实验性纯文本编辑 {#editing}

加载本地源包，选择受支持的形状，然后使用其文本控件。`replacePlainText(pageId, shapeId, text)`、`undo()` 和 `redo()` 通过每个句柄提供相同的操作。通过 `document` 赋值的模型没有包字节，保持只读。

编辑会拒绝与母版关联的文本、富文本、字段、已签名的包以及宏内容。历史记录有上限，可能会丢弃较早的撤销状态。已保存的公式缓存不会被重新计算。尚未实现通用绘图和样式编辑。

`exportVsdx()` 返回原始或已编辑的包字节以及诊断信息，用于明确的副本下载。它绝不会覆盖输入文件。在原生 Visio 中重新打开以及无损往返保真度仍未得到验证。

## 导出静态 SVG {#export}

`viewer.exportSvg()` 返回当前页面的 SVG 字符串、尺寸、字节长度和兼容性诊断。它不会改变选择、缩放和源文档。演示环境中的“导出 SVG”按钮会执行明确的本地下载。

受支持的背景内容和经过验证的栅格图像会被嵌入。字体不会被嵌入；文本换行及其他近似处理会在 SVG 描述和元数据中保持可见。导出上限为 16 MiB，并带有保守的早期资源检查。调用方可以通过 `{ maxBytes }` 降低该上限。这不是可编辑的 VSDX 输出。下游的栅格化程序必须自行限制尺寸和像素面积。

## 准备页面快照 {#prepare-page-snapshots}

`createPrintSnapshot()` 通过 API 返回不可变的 SVG 页面产物和诊断信息。它保留绘图尺寸，但不会应用原生打印机设置、打开打印对话框或实现 PDF 导出。这是为将来的打印工作流准备数据。

## 属性与事件 {#properties-and-events}

| 接口     | 当前约定                                                                                                                                       |
| -------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| 属性     | `document`, `pageIndex`, `zoom`, `showToolbar`, `ribbonAddIns`                                                                                 |
| 事件     | `document-load`, `document-change`, `document-error`, `page-change`, `zoom-change`, `shape-select`, `selection-change`, `office-ribbon-add-in` |
| 生命周期 | `update`, `load`, `fit`, `exportSvg`, `createPrintSnapshot`, `replacePlainText`, `undo`, `redo`, `cancelEdit`, `exportVsdx`, `destroy`         |

请在宿主元素存在之后再挂载，并在框架视图卸载时调用 `destroy()`。重复销毁是安全的。部分更新会保持未提供的属性不变。完整参考见[查看器 API](/zh/api)。

## 采用前的验证 {#verification-before-adoption}

在查看器检出目录中运行 `npm run check` 和 `npm run test:browser`。核心解析器测试属于同级引擎。这些检查只能证明其所断言的行为，并不能证明与 Microsoft Visio 在视觉上一致。
