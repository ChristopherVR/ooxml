import { parseVsdx } from 'visio-core';

const check = (condition, message) => {
	if (!condition) throw new Error(message);
};
const sameBytes = (a, b) => a.length === b.length && a.every((byte, index) => byte === b[index]);

/** Exercise the shared controller through six installed native component handles. */
export async function verifyWorkspaceReplacement(viewer, framework, changes) {
	await viewer.createBlankDrawing({ width: 6, height: 4 });
	const controller = viewer.controller;
	const pageId = controller.state.document.pages[0].id;
	const originalText = ['😀aa a\n', 'a A'];
	await viewer.applyEdits([
		{ type: 'insert-page', pageId: '1', afterPageId: pageId, name: 'Second' },
	]);
	await viewer.applyEdits([
		...[pageId, '1'].map((target, index) => ({
			type: 'create-text-box',
			pageId: target,
			shapeId: '1',
			x: 2,
			y: 2,
			width: 2,
			height: 1,
			text: originalText[index],
		})),
	]);
	viewer.selectAll();
	const before = viewer.exportVsdx().bytes;
	const selected = JSON.stringify(controller.state.selectedShapes);
	const native = [];
	const changed = (event) => native.push(event.detail);
	viewer.element.addEventListener('document-change', changed);
	try {
		let token = controller.captureTextReplaceToken('all-pages');
		const old = token;
		const plan = controller.planTextReplacement(token, {
			query: 'a',
			replacement: '$&',
			mode: 'all',
		});
		check(plan.replacementCount === 4, `${framework}: full literal occurrences`);
		token = await controller.applyTextReplacePlan(plan, token);
		check(
			!controller.isTextReplaceTokenCurrent(old) && controller.isTextReplaceTokenCurrent(token),
			`${framework}: replacement source generation`,
		);
		const state = controller.state;
		check(
			JSON.stringify(state.document.pages.map((page) => page.shapes[0].text.plainText)) ===
				JSON.stringify(['😀$&$& $&\n', '$& A']),
			`${framework}: literal source replacement on both pages`,
		);
		check(JSON.stringify(state.selectedShapes) === selected, `${framework}: Replace All selection`);
		check(
			changes.at(-1).document === state.document && native.at(-1).document === state.document,
			`${framework}: replacement native and binding events`,
		);
		const saved = viewer.exportVsdx().bytes;
		check(
			(await parseVsdx(saved)).pages[0].shapes[0].text.plainText === '😀$&$& $&\n',
			`${framework}: replacement exported source`,
		);
		await viewer.undo();
		check(
			sameBytes(viewer.exportVsdx().bytes, before) &&
				JSON.stringify(controller.state.selectedShapes) === selected,
			`${framework}: byte-exact Replace All undo`,
		);
		await viewer.redo();
		check(sameBytes(viewer.exportVsdx().bytes, saved), `${framework}: byte-exact Replace All redo`);
		token = controller.captureTextReplaceToken('current-page');
		const first = controller.getTextReplaceOccurrences(token, '$&')[0];
		token = controller.selectTextReplaceOccurrence(token, first, '$&');
		const current = controller.planTextReplacement(token, {
			query: '$&',
			replacement: 'Z',
			mode: 'current',
			current: first,
		});
		token = await controller.applyTextReplacePlan(current, token);
		check(
			controller.state.document.pages[0].shapes[0].text.plainText === '😀Z$& $&\n' &&
				controller.getTextReplaceOccurrences(token, '$&').length === 2,
			`${framework}: current replacement and continued owned scope`,
		);
		const unchanged = viewer.exportVsdx().bytes;
		const stale = controller.planTextReplacement(token, {
			query: '$&',
			replacement: 'bad',
			mode: 'all',
		});
		viewer.clearSelection();
		let rejected = false;
		try {
			await controller.applyTextReplacePlan(stale, token);
		} catch {
			rejected = true;
		}
		check(
			rejected && sameBytes(viewer.exportVsdx().bytes, unchanged),
			`${framework}: manual selection invalidates replacement intent`,
		);
	} finally {
		viewer.element.removeEventListener('document-change', changed);
	}
}
