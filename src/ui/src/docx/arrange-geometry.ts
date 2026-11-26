import type { DrawingBox } from 'ooxml-core/docx';
import type { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { boxOf, isFloat } from './print-float-locate';

export function drawingGeometry(core: EditorCore, view: EditorView, pos: number) {
	if (core.pages.viewMode !== 'print' || view !== core.view) return;
	const node = view.state.doc.nodeAt(pos);
	const $pos = view.state.doc.resolve(pos);
	if (!node || $pos.depth === 0 || !isFloat(node)) return;

	const page = core.model.page;
	const margins: DrawingBox = {
		x: page.marginLeft,
		y: page.marginTop,
		width: page.width - page.marginLeft - page.marginRight,
		height: page.height - page.marginTop - page.marginBottom,
	};
	const pageBox: DrawingBox = { x: 0, y: 0, width: page.width, height: page.height };
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
