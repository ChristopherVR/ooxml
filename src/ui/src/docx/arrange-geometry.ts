import type { DrawingBox, PicturePlacement } from 'ooxml-core/docx';
import { floatPosition } from 'ooxml-core/docx/layout';
import type { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { boxOf, isFloat } from './print-float-locate';

export function drawingGeometry(
	core: EditorCore,
	view: EditorView,
	pos: number,
	placement: PicturePlacement,
) {
	const node = view.state.doc.nodeAt(pos);
	const $pos = view.state.doc.resolve(pos);
	if (!node || $pos.depth === 0) return;
	const widthPx = Number(node.attrs.widthPx);
	const heightPx = Number(node.attrs.heightPx);
	const paper = core.shell.paper;
	const paragraph = view.nodeDOM($pos.before());
	const page = core.model.page;
	const margins: DrawingBox = {
		x: page.marginLeft,
		y: page.marginTop,
		width: page.width - page.marginLeft - page.marginRight,
		height: page.height - page.marginTop - page.marginBottom,
	};
	const pageBox: DrawingBox = { x: 0, y: 0, width: page.width, height: page.height };
	if (core.pages.viewMode === 'print' && view === core.view && isFloat(node)) {
		const pages = core.shell.printLayout?.element;
		const id = CSS.escape(String($pos.parent.attrs.id));
		let index = 0;
		$pos.parent.forEach((child, offset) => {
			if ($pos.start() + offset < pos && isFloat(child)) index++;
		});
		const rendered = pages?.querySelector<HTMLElement>(
			`.dve-print-float[data-block-id="${id}"][data-float-index="${index}"]`,
		);
		const sheet = rendered?.closest<HTMLElement>('.dve-print-page');
		if (!rendered || !sheet) return;
		const line = sheet.querySelector<HTMLElement>(`.dve-print-line[data-block-id="${id}"]`);
		const scale = sheet.offsetHeight
			? sheet.getBoundingClientRect().height / sheet.offsetHeight || 1
			: 1;
		const paragraphY = line
			? (line.getBoundingClientRect().top - sheet.getBoundingClientRect().top) / scale
			: Number(rendered.dataset.originY);
		const origins = { x: Number(rendered.dataset.originX), y: Number(rendered.dataset.originY) };
		return { box: boxOf(rendered), origins, paragraphY, margins, pageBox, sheet };
	}
	if (view !== core.view || !(paragraph instanceof HTMLElement) || !paper) return;
	const scale = paper.offsetHeight
		? paper.getBoundingClientRect().height / paper.offsetHeight || 1
		: 1;
	const paragraphY =
		(paragraph.getBoundingClientRect().top - paper.getBoundingClientRect().top) / scale;
	const { xPx: x, yPx: y } = floatPosition(
		{
			widthPx,
			heightPx,
			...(placement.relativeFrom ? { relativeFromH: placement.relativeFrom } : {}),
			...(placement.align ? { alignH: placement.align } : {}),
			...(placement.offsetXPx !== undefined ? { offsetXPx: placement.offsetXPx } : {}),
			...(placement.relativeFromV ? { relativeFromV: placement.relativeFromV } : {}),
			...(placement.alignV ? { alignV: placement.alignV } : {}),
			...(placement.offsetYPx !== undefined ? { offsetYPx: placement.offsetYPx } : {}),
		},
		{
			widthPx: page.width,
			heightPx: page.height,
			marginTopPx: page.marginTop,
			marginBottomPx: page.marginBottom,
			marginLeftPx: page.marginLeft,
			marginRightPx: page.marginRight,
			index: 0,
			sectionIndex: 0,
			pageInSection: 0,
			columns: [],
		},
		{ xPx: 0, widthPx: margins.width },
		{ topPx: paragraphY, heightPx: paragraph.offsetHeight },
	);
	return {
		box: { x, y, width: widthPx, height: heightPx },
		origins: { x: 0, y: 0 },
		paragraphY,
		margins,
		pageBox,
		sheet: paper,
	};
}
