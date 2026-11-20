## Added

- `ooxml-core/docx`: `PicturePlacement.layoutInCell` is `false` for a floating object written with `layoutInCell="0"`, and a floating chart's `InlineImage` now carries its `placement` (wrapping, frames, alignment and offsets) as floating pictures and shapes do.
- `ooxml-core/docx/layout`: a paragraph box in a table cell lists the floating objects laid out in that cell (`LayoutParagraphBox.floats`), and a page's float box marks them with `inCell`.

## Fixed

- `ooxml-core/docx/layout`, `ooxml-ui` and Word viewers: floating pictures, shapes, text boxes and charts anchored in a table cell are drawn in Print Layout where Word draws them. Their column, margin, page and character frames and their vertical frames are taken from the cell, the cell's text wraps around a square, tight or top-and-bottom object, and the row grows to hold it, as in Word 2013 and later, which lay every such object out in its cell. In documents saved in an earlier compatibility mode, an object with "Layout in table cell" turned off is placed on the page instead. These objects used to be left off the page.
- `ooxml-core/docx/layout`, `ooxml-ui` and Word viewers: floating charts are drawn in Print Layout, in the body and in table cells, painted like inline charts and positioned and wrapped like floating pictures, instead of being left off the page. A floating chart can be selected, moved and resized on the page, and saving keeps its new position.
- `ooxml-core/docx/layout`: a floating object positioned relative to the character or line it is anchored to is measured from that character and line, as Word does, instead of from the column and the paragraph's top.
