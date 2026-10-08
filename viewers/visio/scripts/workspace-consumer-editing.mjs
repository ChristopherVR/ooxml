// Every operation runs through actual installed native handles and their shared workers.
import { parseVsdx } from 'visio-core';

const check = (condition, message) => {
	if (!condition) throw new Error(message);
};
export async function verifyWorkspaceEditing(viewer, framework, events) {
	const { selections, primary, changes, nativeSelections } = events;
	await viewer.createBlankDrawing({ width: 6, height: 4 });
	const blank = viewer.controller.state;
	check(
		blank.document.pages.length === 1 &&
			blank.document.pages[0].width === 6 &&
			blank.document.pages[0].height === 4 &&
			blank.document.pages[0].shapes.length === 0,
		`${framework}: blank source dimensions`,
	);
	check(
		blank.selectedShapes.length === 0 &&
			blank.edit.sourceAvailable &&
			!blank.edit.dirty &&
			!blank.edit.canUndo &&
			!blank.edit.canRedo,
		`${framework}: clean editable new source`,
	);
	check(viewer.element.fileName === 'New drawing.vsdx', `${framework}: New filename`);
	const blankPageId = blank.document.pages[0].id;
	await viewer.applyEdits([
		{
			type: 'create-rectangle',
			pageId: blankPageId,
			shapeId: '1',
			x: 2,
			y: 2,
			width: 1,
			height: 1,
		},
	]);
	check(
		viewer.controller.state.edit.dirty &&
			viewer.controller.state.document.pages[0].shapes.length === 1,
		`${framework}: new source edit`,
	);
	const originalShape = viewer.controller.state.document.pages[0].shapes[0];
	viewer.selectShapes([{ id: originalShape.id, name: originalShape.name, pageId: blankPageId }]);
	await verifyPaint(viewer, framework, changes, blankPageId, originalShape.id);
	const beforeResize = viewer.exportVsdx().bytes;
	const beforeSelection = JSON.stringify(viewer.controller.state.selectedShapes);
	await viewer.applyEdits([
		{
			type: 'resize-shape',
			pageId: blankPageId,
			shapeId: originalShape.id,
			width: 2,
			height: 1,
			anchor: { x: 0, y: 0 },
		},
	]);
	const resized = viewer.controller.state.document.pages[0].shapes[0];
	check(
		resized.width === 2 &&
			resized.height === 1 &&
			resized.rotation.pinX === 2.5 &&
			resized.rotation.pinY === 2,
		`${framework}: fixed opposite anchor resize`,
	);
	check(
		JSON.stringify(viewer.controller.state.selectedShapes) === beforeSelection &&
			viewer.controller.state.selectedShape === viewer.controller.state.selectedShapes[0],
		`${framework}: anchored resize selection`,
	);
	await viewer.undo();
	const restored = viewer.exportVsdx().bytes;
	check(
		restored.length === beforeResize.length &&
			restored.every((byte, index) => byte === beforeResize[index]) &&
			JSON.stringify(viewer.controller.state.selectedShapes) === beforeSelection,
		`${framework}: byte-exact anchored resize undo`,
	);
	await viewer.undo();
	check(
		!viewer.controller.state.edit.dirty &&
			viewer.controller.state.document.pages[0].shapes.length === 0,
		`${framework}: clean new source undo`,
	);
	const beforeText = viewer.exportVsdx().bytes,
		text = `Installed ${framework}\nBlank paragraph\n`;
	await viewer.applyEdits([
		{
			type: 'create-text-box',
			pageId: blankPageId,
			shapeId: '2',
			x: 2,
			y: 2,
			width: 2,
			height: 1,
			text,
		},
	]);
	const textState = viewer.controller.state,
		textShape = textState.document.pages[0].shapes[0];
	check(
		textShape.text.plainText === text &&
			textShape.style.fill === 'none' &&
			textShape.style.linePattern === 0,
		`${framework}: logical text and paint-free text box`,
	);
	check(
		changes.at(-1).document === textState.document &&
			changes.at(-1).kind === 'edit' &&
			textState.edit.dirty,
		`${framework}: created model through binding callback`,
	);
	viewer.selectAll();
	check(
		selections.at(-1) === viewer.controller.state.selectedShapes &&
			nativeSelections.at(-1) === selections.at(-1) &&
			primary.at(-1)?.id === textShape.id,
		`${framework}: created text selection events`,
	);
	await viewer.undo();
	const beforeTextRestored = viewer.exportVsdx().bytes;
	check(
		beforeTextRestored.length === beforeText.length &&
			beforeTextRestored.every((byte, index) => byte === beforeText[index]) &&
			viewer.controller.state.document.pages[0].shapes.length === 0 &&
			viewer.controller.state.selectedShapes.length === 0,
		`${framework}: byte-exact text creation undo`,
	);
}

async function verifyPaint(viewer, framework, changes, pageId, shapeId) {
	const native = [];
	const changed = (event) => native.push(event.detail);
	viewer.element.addEventListener('document-change', changed);
	const selected = JSON.stringify(viewer.controller.state.selectedShapes);
	let beforeNoLine;
	const patches = [
		{ linePattern: 0 },
		{ linePattern: 23, lineTransparency: 33.26 },
		...Array.from({ length: 25 }, (_, fillPattern) => ({
			fillPattern,
			fillColor: '#123456',
			fillBackgroundColor: '#abcdef',
			fillTransparency: 12.26,
		})),
	];
	try {
		for (const patch of patches) {
			const before = viewer.exportVsdx().bytes;
			await viewer.applyEdits([{ type: 'format-shape', pageId, shapeId, ...patch }]);
			const state = viewer.controller.state,
				style = state.document.pages[0].shapes[0].style;
			if (patch.linePattern !== undefined)
				check(style.linePattern === patch.linePattern, `${framework}: saved line pattern`);
			if (patch.lineTransparency !== undefined)
				check(
					Math.abs(style.lineColorOpacity - 0.665) < 1e-12,
					`${framework}: half-percent line transparency`,
				);
			if (patch.fillPattern !== undefined) {
				check(
					style.fillPatternIndex === patch.fillPattern && style.fillBackgroundColor === '#abcdef',
					`${framework}: saved fill pattern/background`,
				);
				check(
					style.fillForegroundOpacity === 0.875 && style.fillBackgroundOpacity === 0.875,
					`${framework}: half-percent foreground/background transparency`,
				);
				check(
					patch.fillPattern === 0 ? style.fill === 'none' : style.fill === '#123456',
					`${framework}: explicit fill pattern precedence`,
				);
			}
			check(
				changes.at(-1).document === state.document && native.at(-1).document === state.document,
				`${framework}: paint model through native and binding events`,
			);
			check(
				JSON.stringify(state.selectedShapes) === selected &&
					state.selectedShape === state.selectedShapes[0],
				`${framework}: paint preserves selection`,
			);
			const saved = viewer.exportVsdx().bytes;
			check(!sameBytes(saved, before), `${framework}: paint changes retained source`);
			check(
				JSON.stringify((await parseVsdx(saved)).pages[0].shapes[0].style) === JSON.stringify(style),
				`${framework}: paint export/reparse`,
			);
			if (patch.linePattern === 0) {
				beforeNoLine = before;
				continue;
			}
			await viewer.undo();
			check(
				sameBytes(viewer.exportVsdx().bytes, before) &&
					JSON.stringify(viewer.controller.state.selectedShapes) === selected,
				`${framework}: byte-exact paint undo`,
			);
			if (beforeNoLine) {
				await viewer.undo();
				check(
					sameBytes(viewer.exportVsdx().bytes, beforeNoLine),
					`${framework}: byte-exact No Line undo`,
				);
				beforeNoLine = undefined;
			}
		}
	} finally {
		viewer.element.removeEventListener('document-change', changed);
	}
}

const sameBytes = (left, right) =>
	left.length === right.length && left.every((byte, index) => byte === right[index]);
