/** The engine loads on first use, so the chat can paint before any Office code arrives. */
const automation = () => import('ooxml-core/automation');

/** Keeps items from the front of `list` while their JSON fits in `budget` characters. */
function fit(list, budget) {
	let used = 0,
		count = 0;
	for (const item of list) {
		used += JSON.stringify(item).length + 1;
		if (used > budget) break;
		count++;
	}
	return count;
}

/**
 * A file as the assistant reads it, within `budget` JSON characters. Nothing is cut silently: when
 * the budget runs out the view says what it left out (`omitted`), so the model can say so too.
 */
export async function inspect(doc, budget = 150000) {
	const { inspectDocx, readXlsxFilledCells, inspectVisio } = await automation();
	if (doc.kind === 'docx') {
		const source = await inspectDocx(doc.bytes);
		const paragraphs = source.blocks
			.flatMap((b) =>
				b.type === 'paragraph'
					? [b]
					: b.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
			)
			.map((p) => ({
				id: p.id,
				runs: p.runs.map((r, runIndex) => ({ runIndex, text: r.text })),
			}));
		const kept = fit(paragraphs, budget);
		return {
			paragraphs: paragraphs.slice(0, kept),
			...(kept < paragraphs.length
				? { omitted: `Paragraphs ${kept + 1} to ${paragraphs.length} were left out to fit.` }
				: {}),
		};
	}
	if (doc.kind === 'xlsx') {
		// A filled cell is about 60 JSON characters; read as many as the budget holds.
		const book = await readXlsxFilledCells(doc.bytes, { maxCells: Math.floor(budget / 60) });
		const left = book.sheets.filter((s) => s.omitted);
		return {
			sheets: book.sheets.map(({ index, name, cells }) => ({ index, name, cells })),
			...(left.length
				? {
						omitted: `Left out to fit: ${left.map((s) => `${s.omitted} of ${s.total} filled cells on ${s.name}`).join('; ')}.`,
					}
				: {}),
		};
	}
	const view =
		doc.kind === 'vsdx'
			? await inspectVisio(doc.bytes)
			: await (await import('ooxml-core/pptx')).inspectPresentationText(doc.bytes);
	const list = doc.kind === 'vsdx' ? view.pages : view.slides;
	const kept = fit(list ?? [], budget);
	if (!list || kept >= list.length) return view;
	const unit = doc.kind === 'vsdx' ? 'Pages' : 'Slides';
	return {
		...view,
		[doc.kind === 'vsdx' ? 'pages' : 'slides']: list.slice(0, kept),
		omitted: `${unit} ${kept + 1} to ${list.length} were left out to fit.`,
	};
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
/** The core's workbook edits for the `cells` changes to one file, in order. */
const cellEdits = (changes, documentId) =>
	changes
		.filter((c) => c.type === 'cells' && c.documentId === documentId)
		.map((c) => ({ op: 'setCells', sheet: c.sheetIndex, cells: c.cells }));

const shown = (value) =>
	value === null || value === undefined
		? '(empty)'
		: typeof value === 'object'
			? value.error
			: typeof value === 'boolean'
				? String(value).toUpperCase()
				: String(value);

/**
 * What the `cells` changes to spreadsheet `doc` would change, compared by the core before anything
 * is saved. `effects` are the differences beyond the cells typed into (formulas that recalculate,
 * spill ranges), at most `limit`, and `more` counts the rest. `proposal` is what
 * {@link prepareChanges} applies when exactly these changes are approved.
 */
export async function reviewSpreadsheet(changes, doc, limit = 50) {
	const edits = cellEdits(changes, doc.id);
	if (!edits.length) return { effects: [], more: 0 };
	const [{ proposeXlsxEdits }, { formatAddress }] = await Promise.all([
		automation(),
		import('ooxml-core/xlsx'),
	]);
	const { proposal, comparison } = await proposeXlsxEdits(doc.bytes, edits);
	const typed = new Set(
		edits.flatMap((e) => e.cells.map((cell) => `${e.sheet}!${cell.address.toUpperCase()}`)),
	);
	const effects = [];
	for (const sheet of comparison.sheets)
		for (const cell of sheet.cells) {
			const address = formatAddress(cell.address);
			if (cell.kind === 'format' || typed.has(`${sheet.afterIndex}!${address}`)) continue;
			effects.push({
				where: sheet.name,
				address,
				before: shown(cell.before.value),
				after: shown(cell.after.value),
			});
		}
	return {
		proposal,
		effects: effects.slice(0, limit),
		more: Math.max(0, effects.length - limit) + (comparison.truncated ? 1 : 0),
	};
}

/**
 * The files with `changes` applied, as new revisions. Spreadsheet changes go through the core's
 * proposal: one reviewed for exactly these changes (`proposals`, by file id) is applied only to the
 * bytes it was reviewed against, so what is saved is what was shown; otherwise they are proposed
 * and applied in one step.
 */
export async function prepareChanges(changes, docs, proposals = {}) {
	const { setDocxRunText, editVisio, applyXlsxProposal, proposeXlsxEdits } = await automation();
	validateChanges(changes, docs);
	const updates = new Map();
	for (const original of docs) {
		const edits = cellEdits(changes, original.id);
		if (!edits.length) continue;
		const reviewed = proposals[original.id];
		const proposal =
			reviewed && JSON.stringify(reviewed.edits) === JSON.stringify(edits)
				? reviewed
				: (await proposeXlsxEdits(original.bytes, edits)).proposal;
		const { bytes } = await applyXlsxProposal(original.bytes, proposal);
		updates.set(original.id, {
			...original,
			bytes,
			revision: original.revision + 1,
			modified: Date.now(),
		});
	}
	for (const c of changes) {
		if (c.type === 'cells') continue;
		const original = docs.find((d) => d.id === c.documentId);
		const doc = updates.get(c.documentId) ?? original;
		let bytes;
		if (c.type === 'word_run')
			bytes = (await setDocxRunText(doc.bytes, c.paragraphId, c.runIndex, c.text)).bytes;
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
