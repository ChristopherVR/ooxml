// Synthetic browser transport: this test never accesses the operating system clipboard.
export async function verifyWorkspaceClipboard(viewer, framework) {
	const check = (value, message) => {
		if (!value) throw new Error(`${framework}: ${message}`);
	};
	const descriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
	let text = '',
		writes = 0,
		reads = 0,
		failWrite = false;
	Object.defineProperty(navigator, 'clipboard', {
		configurable: true,
		value: {
			async writeText(value) {
				if (failWrite) throw new Error('Synthetic clipboard denied');
				text = value;
				writes++;
			},
			async readText() {
				reads++;
				return text;
			},
		},
	});
	const ready = async () => {
		for (let attempt = 0; attempt < 200; attempt++) {
			const state = viewer.controller.state.clipboard;
			if (state.error) throw state.error;
			if (state.ready) return;
			await new Promise((done) => setTimeout(done, 10));
		}
		throw new Error(`${framework}: clipboard preparation timed out`);
	};
	try {
		await ready();
		const originals = viewer.controller.state.selectedShapes.map((shape) => shape.id);
		const count = viewer.controller.state.document.pages[0].shapes.length;
		await viewer.copySelection();
		check(writes === 1 && text.startsWith('OOXML-VISIO-SHAPES/1\n'), 'portable clipboard write');
		check(
			viewer.controller.state.document.pages[0].shapes.length === count,
			'Copy preserves source',
		);
		const pageId = viewer.controller.state.document.pages[0].id;
		const originalText = viewer.controller.state.document.pages[0].shapes.find(
			(shape) => shape.id === originals[0],
		).text?.plainText;
		await viewer.replacePlainText(pageId, originals[0], 'Changed after copy');
		await viewer.pasteSelection();
		check(reads === 1, 'Paste reads actual transport');
		const pasted = viewer.controller.state.selectedShapes.map((shape) => shape.id);
		check(
			pasted.length === 2 && viewer.controller.state.document.pages[0].shapes.length === count + 2,
			'source-backed Paste',
		);
		check(
			viewer.controller.state.document.pages[0].shapes.find((shape) => shape.id === pasted[0]).text
				?.plainText === originalText,
			'old clipboard source preserved',
		);
		await viewer.undo();
		check(
			JSON.stringify(viewer.controller.state.selectedShapes.map((shape) => shape.id)) ===
				JSON.stringify(originals),
			'Paste undo selection',
		);
		await viewer.redo();
		await ready();
		failWrite = true;
		let refused = false;
		try {
			await viewer.cutSelection();
		} catch {
			refused = true;
		}
		check(
			refused && viewer.controller.state.document.pages[0].shapes.length === count + 2,
			'failed write never deletes',
		);
		failWrite = false;
		await viewer.cutSelection();
		check(
			writes === 2 &&
				viewer.controller.state.document.pages[0].shapes.length === count &&
				viewer.controller.state.selectedShapes.length === 0,
			'Cut writes before atomic delete',
		);
		await viewer.undo();
		check(
			JSON.stringify(viewer.controller.state.selectedShapes.map((shape) => shape.id)) ===
				JSON.stringify(pasted),
			'Cut undo restores selection',
		);
		await viewer.redo();
		await viewer.pasteSelection();
		check(
			reads === 2 && viewer.controller.state.selectedShapes.length === 2,
			'empty selection permits Paste',
		);
		text = 'Unrelated synthetic clipboard text';
		refused = false;
		try {
			await viewer.pasteSelection();
		} catch {
			refused = true;
		}
		check(
			refused &&
				reads === 3 &&
				viewer.controller.state.document.pages[0].shapes.length === count + 2,
			'invalid actual clipboard never uses cached payload',
		);
	} finally {
		if (descriptor) Object.defineProperty(navigator, 'clipboard', descriptor);
		else delete navigator.clipboard;
	}
}
