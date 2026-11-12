# 文档模型中的度量单位

`DocumentModel` 以 WordprocessingML 在磁盘上使用的单位存储度量值。自单位品牌类型（unit brands）引入以来，这些单位已成为 TypeScript 类型的一部分：在需要长度的地方传入字号，或在需要整数缇（twips）的地方传入未取整的像素值，都会在编译时报错，而不会生成损坏的文件。

品牌类型仅在编译时生效。`Twips` 在运行时就是普通的 JavaScript 数字，并且仍可赋值给 `number`，因此算术运算、比较和 JSON 序列化都不受影响。

## 品牌类型 {#brands}

所有品牌类型都从 `docx-core` 导出。

| 类型           | 单位                       | 架构类型                | 构造函数（严格）  | 取整构造函数           |
| -------------- | -------------------------- | ----------------------- | ----------------- | ---------------------- |
| `Twips`        | 1/20 磅，非负              | `ST_TwipsMeasure`       | `twips(n)`        | `roundTwips(n)`        |
| `SignedTwips`  | 1/20 磅，可为负            | `ST_SignedTwipsMeasure` | `signedTwips(n)`  | `roundSignedTwips(n)`  |
| `HalfPoints`   | 1/2 磅（`w:sz` 字号）      | `ST_HpsMeasure`         | `halfPoints(n)`   | `roundHalfPoints(n)`   |
| `EighthPoints` | 1/8 磅（边框宽度）         | `ST_EighthPointMeasure` | `eighthPoints(n)` | `roundEighthPoints(n)` |
| `Emu`          | 1/914400 英寸（DrawingML） | `ST_Coordinate`         | `emu(n)`          | `roundEmu(n)`          |

`Twips` 是 `SignedTwips` 的子类型：无符号度量可以赋值给有符号字段，反之则不行。

**严格构造函数**在数值不是整数（且是安全整数）时抛出 `RangeError`；对于 `twips`，负值同样会抛出。解析器和测试使用这些构造函数。

**取整构造函数**用于计算得出的值（标尺拖动、缩放计算、像素换算）。它们采用四舍五入（远离零方向）的方式取整，仅在输入为 `NaN` 或无穷大时抛出 `RangeError`。`roundTwips` 和 `roundEighthPoints` 会将负结果钳制为零。

**从现实单位转换**时，会四舍五入到最接近的整数目标单位：`twipsFromPoints`、`signedTwipsFromPoints`、`twipsFromPixels`（每个 CSS 像素 15 缇）、`signedTwipsFromPixels`、`twipsFromInches`、`halfPointsFromPoints`、`eighthPointsFromPoints`、`emuFromPixels`。

**转换为现实单位**时返回普通数字：`twipsToPoints`、`twipsToPixels`、`twipsToInches`、`halfPointsToPoints`、`eighthPointsToPoints`、`emuToPixels`。

`units.ts` 之外不应使用 `as Twips` 这种类型断言；请使用构造函数。

## 各模型字段使用的品牌类型 {#which-model-fields-carry-which-brand}

字段名称保持不变。

- **段落**（`Paragraph`、`ParagraphFormatting`、样式定义）：`spacingBeforeTwips`、`spacingAfterTwips`、`firstLineTwips`、`hangingTwips` 为 `Twips`；`lineSpacingTwips`、`indentLeftTwips`、`indentRightTwips`、`indentStartTwips`、`indentEndTwips` 为 `SignedTwips`。`TabStop.posTwips` 为 `SignedTwips`。
- **编号级别与列表标签**：`indentLeftTwips` 为 `SignedTwips`；`hangingTwips` 和 `firstLineTwips` 为 `Twips`。
- **节**：`pageWidthTwips`、`pageHeightTwips`、`marginLeftTwips`、`marginRightTwips`、`headerDistanceTwips`、`footerDistanceTwips`、`gutterTwips`，以及分栏的 `widthTwips`/`spacingTwips` 为 `Twips`；`marginTopTwips` 和 `marginBottomTwips` 为 `SignedTwips`（Word 允许上、下页边距为负值）。
- **表格**：`Table.grid` 为 `Twips[]`；`Table.widthTwips`、`TableCell.widthTwips` 和 `TableRowProperties.heightTwips` 为 `Twips`；`Table.indentTwips` 为 `SignedTwips`（Word 经常写入负的表格缩进）；`TableCellMargins.top/bottom/left/right` 为 `SignedTwips`。
- **边框**：`TableBorderSide.sizeEighthPoints` 为 `EighthPoints`。
- **文本串**：`TextRun.characterSpacingTwips` 为 `SignedTwips`。
- **排版输入**（`ooxml-core/docx/layout`）：`LayoutParagraph` 上的段落 `*Twips` 字段与模型使用相同的品牌类型。

### 刻意不加品牌类型的字段 {#fields-that-are-deliberately-not-branded}

- `TextRun.fontSize` 以**磅**为单位，可以是小数（`10.5`）。`HalfPoints` 是磁盘上的 `w:sz` 值；解析器会进行转换。写回字号时请使用 `halfPointsFromPoints`。
- `TableBorderSide.spacePoints` 是整数**磅**（`ST_PointMeasure`）。
- `widthPx`、`heightPx`、`offsetXPx`、`offsetYPx` 以及排版中所有的 `*Px` 字段，都是 96 dpi 下的 **CSS 像素**。它们是派生的设备空间值，且经常是小数，因此保持为普通数字。图片尺寸以像素存储（由 EMU 派生），而不是 `Emu`。
- `DocumentModel.page`（`width`、`height`、页边距）以 CSS 像素表示；`SectionProperties` 才是权威的缇数来源。

## 构建模型值 {#building-model-values}

模型字段接受带品牌的数字，因此构建或编辑模型对象的代码需要通过构造函数来创建它们：

```ts
import { twips, eighthPoints, roundSignedTwips, twipsFromPixels } from 'docx-core';
paragraph.spacingAfterTwips = twips(240);
table.grid = [twips(3000), twips(3000)];
side.sizeEighthPoints = eighthPoints(8);

// Computed values: round explicitly at the boundary
paragraph.indentLeftTwips = roundSignedTwips(dragPx * 15);
section.pageWidthTwips = twipsFromPixels(page.width);
```

读取不受影响：带品牌的值仍是数字，因此 `paragraph.spacingAfterTwips ?? 0` 和 `twipsToPixels(x)` 照常工作。当你需要把一个你信任其为整数的来源（例如从你自己的 JSON 中读回的值）作为普通数字放入带品牌的槽位时，请先通过 `twips()` 或 `signedTwips()` 处理，这样错误的值会立即报错。`twips(n)` 会拒绝负数；有符号的值请使用 `signedTwips(n)`。

## 解析遵循架构的符号性 {#parsing-follows-the-schema-signedness}

由于模型现在区分有符号和无符号长度，解析器会按属性应用架构中的类型。以下属性的负值会被拒绝（字段保持未定义）：`w:spacing/@before|@after`、`w:ind/@firstLine|@hanging`、页面的 `w:pgMar/@left|@right|@header|@footer|@gutter`、表格网格、宽度和行高。而以下属性的负值会被保留：`w:ind/@left|@right|@start|@end`、`w:spacing/@line`、`w:tab/@pos`、`w:pgMar/@top|@bottom`、`w:tblInd`，以及 `w:spacing`（字符间距）。`validateModel`（保存前的验证器）采用相同的分类方式，并拒绝非整数。

## 编辑器属性边界 {#editor-attribute-boundary}

ProseMirror 节点属性没有类型。编辑器在一个地方将它们转换为模型（`web-component` 中的 `attr-units.ts`）：数值会被取整为整数缇；非数字、非有限或（对于无符号字段）为负的值会被丢弃，而不会被写入。
