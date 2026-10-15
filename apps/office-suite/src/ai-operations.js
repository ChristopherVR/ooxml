/** The engine loads on first use, so the chat can paint before any Office code arrives. */
const automation = () => import('ooxml-core/automation');

export async function inspect(doc) {
	const { inspectDocx, inspectXlsx, readXlsxRange, inspectVisio } = await automation();
	if (doc.kind === 'docx') {
		const source = await inspectDocx(doc.bytes);
		const paragraphs = source.blocks.flatMap((b) =>
			b.type === 'paragraph'
				? [b]
				: b.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
		);
		return {
			paragraphs: paragraphs.slice(0, 500).map((p) => ({
				id: p.id,
				runs: p.runs.map((r, runIndex) => ({ runIndex, text: r.text })),
			})),
			truncated: paragraphs.length > 500,
		};
	}
	if (doc.kind === 'xlsx') {
		const book = await inspectXlsx(doc.bytes);
		const sheets = [];
		for (const sheet of book.sheets.slice(0, 4))
			sheets.push({
				...sheet,
				range: 'A1:AD100',
				...(await readXlsxRange(doc.bytes, sheet.index, 'A1:AD100')),
			});
		return {
			sheets: sheets.map((s) => ({
				...s,
				cells: s.cells.filter((c) => c.value !== null || c.formula),
			})),
			scope: 'First four sheets, A1:AD100 only',
		};
	}
	if (doc.kind === 'vsdx') return inspectVisio(doc.bytes);
	const { inspectPresentationText } = await import('ooxml-core/pptx');
	return inspectPresentationText(doc.bytes);
}

// Functions that fetch data, open links or load code when a formula is evaluated.
const BLOCKED_FUNCTION =
	/(?:^|[^A-Z0-9_.])(?:_XL[A-Z]*\.)?(?:WEBSERVICE|FILTERXML|HYPERLINK|IMAGE|RTD|CALL|REGISTER(?:\.ID)?)\s*\(/i;
// Links, UNC shares and drive paths. Matched anywhere, including inside string literals, which
// is deliberately conservative because the formula is not tokenised here.
const EXTERNAL_TARGET = /http|www\.|file:|ftp:|\\\\|[a-z]:\\/i;

/** Returns why a proposed cell input is unsafe, or null when it is allowed. */
export function formulaProblem(input) {
	if (!/^\s*=/.test(input)) return null;
	const formula = input.trim();
	if (BLOCKED_FUNCTION.test(formula))
		return 'it calls a function that can fetch data, open links or run code';
	if (formula.includes('|')) return 'DDE pipe references are not allowed';
	if (EXTERNAL_TARGET.test(formula)) return 'it refers to a link or a file path';
	for (const match of formula.matchAll(/\[([^\]]*)/g))
		if (/[.\\/:]|^\s*\d+\s*$/.test(match[1])) return 'it refers to an external workbook';
	return null;
}

export function validateChanges(changes, docs) {
	if (!Array.isArray(changes) || changes.length > 30)
		throw new Error('The response has too many changes. Ask for a smaller edit.');
	for (const c of changes) {
		const doc = docs.find((d) => d.id === c?.documentId);
		if (!doc || doc.parent)
			throw new Error(
				'Edits must target selected top-level documents. Edit embedded files in their editor.',
			);
		if (
			{ word_run: 'docx', cells: 'xlsx', slide_run: 'pptx', shape_text: 'vsdx' }[c.type] !==
			doc.kind
		)
			throw new Error('The proposed operation does not match the file type.');
		if (c.type !== 'cells' && (typeof c.text !== 'string' || c.text.length > 20000))
			throw new Error('Invalid replacement text.');
		for (const key of c.type === 'word_run'
			? ['paragraphId']
			: c.type === 'slide_run'
				? ['elementId']
				: c.type === 'shape_text'
					? ['pageId', 'shapeId']
					: [])
			if (typeof c[key] !== 'string' || c[key].length > 200)
				throw new Error('Invalid edit target.');
		for (const key of c.type === 'cells'
			? ['sheetIndex']
			: c.type === 'slide_run'
				? ['slideIndex', 'runIndex']
				: c.type === 'word_run'
					? ['runIndex']
					: [])
			if (!Number.isSafeInteger(c[key]) || c[key] < 0) throw new Error('Invalid edit index.');
		if (
			c.type === 'cells' &&
			(!Array.isArray(c.cells) ||
				c.cells.length > 500 ||
				c.cells.some(
					(cell) =>
						typeof cell.address !== 'string' ||
						typeof cell.input !== 'string' ||
						cell.input.length > 20000,
				))
		)
			throw new Error('Invalid cell edits.');
		if (c.type === 'cells')
			for (const cell of c.cells) {
				const problem = formulaProblem(cell.input);
				if (problem)
					throw new Error(
						`The proposed formula in ${cell.address} was rejected because ${problem}. Ask again without it.`,
					);
			}
	}
}
export async function prepareChanges(changes, docs) {
	const { setDocxRunText, setXlsxCells, editVisio } = await automation();
	validateChanges(changes, docs);
	const updates = new Map();
	for (const c of changes) {
		const original = docs.find((d) => d.id === c.documentId);
		const doc = updates.get(c.documentId) ?? original;
		let bytes;
		if (c.type === 'word_run')
			bytes = (await setDocxRunText(doc.bytes, c.paragraphId, c.runIndex, c.text)).bytes;
		if (c.type === 'cells') bytes = (await setXlsxCells(doc.bytes, c.sheetIndex, c.cells)).bytes;
		if (c.type === 'shape_text')
			bytes = (
				await editVisio(doc.bytes, [
					{ type: 'replace-plain-text', pageId: c.pageId, shapeId: c.shapeId, text: c.text },
				])
			).bytes;
		if (c.type === 'slide_run')
			bytes = await (
				await import('ooxml-core/pptx')
			).setPresentationRunText(doc.bytes, c.slideIndex, c.elementId, c.runIndex, c.text);
		updates.set(doc.id, { ...doc, bytes, revision: original.revision + 1, modified: Date.now() });
	}
	return [...updates.values()];
}

/** The bytes of a new file the assistant proposed (see `validateCreate` in ai/protocol.js). */
export async function buildFile(item) {
	const { createDocxFromOutline, createXlsxFromSheets } = await automation();
	if (item.type === 'docx') return createDocxFromOutline(item.blocks);
	if (item.type === 'xlsx') return createXlsxFromSheets(item.sheets);
	const { createPresentationFromOutline } = await import('ooxml-core/pptx');
	return createPresentationFromOutline(item.slides, { title: item.name });
}
