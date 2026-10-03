import {
	analyzeVisioFormula,
	parseVisioFormula,
	type VisioFormulaAst,
	type VisioFormulaReference,
} from './formula.js';

/** Locality proof only, never numeric evaluation or permission to overwrite a redirect. */
export function analyzeVisioMasterFormula(source: string, options: { textFieldFree: boolean }) {
	const ast = parseVisioFormula(source);
	const analysis = analyzeVisioFormula(ast);
	const references = new Map<string, VisioFormulaReference>();
	for (const ref of analysis.references)
		references.set(`${ref.shapeId ?? ''}!${ref.cell.toLowerCase()}`, ref);
	let dynamic = false,
		readsText = false,
		readsTheme = false;
	const queue: VisioFormulaAst[] = [ast];
	while (queue.length) {
		const node = queue.pop()!;
		if (node.kind === 'unary') queue.push(node.operand);
		if (node.kind === 'binary') queue.push(node.left, node.right);
		if (node.kind !== 'call') continue;
		queue.push(...node.args);
		if (['THEME', 'THEMEVAL', 'ISTHEMED'].includes(node.name)) readsTheme = true;
		let local = !analyzeVisioFormula({ kind: 'call', name: node.name, args: [] }).dynamic;
		if (node.name === 'SETATREF')
			local = node.args.length === 1 && node.args[0]!.kind === 'reference';
		if (['TEXTWIDTH', 'TEXTHEIGHT', 'SHAPETEXT'].includes(node.name)) {
			const text = node.args[0];
			const count = node.name === 'TEXTHEIGHT' ? node.args.length === 2 : node.args.length === 1;
			local =
				options.textFieldFree &&
				count &&
				text?.kind === 'reference' &&
				text.reference.shapeId === undefined &&
				text.reference.cell.toLowerCase() === 'thetext';
			readsText = true;
		}
		if (['LUM', 'HUE', 'SAT'].includes(node.name)) {
			local = node.args.length === 1;
			const arg = node.args[0];
			if (arg?.kind === 'string') {
				// Conservatively treat possible string cell names as implicit reads.
				local = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(arg.value);
				if (local) references.set(`!${arg.value.toLowerCase()}`, { cell: arg.value });
			}
		}
		if (node.name === 'STRSAME') local = node.args.length === 2;
		if (!local) dynamic = true;
	}
	return { ...analysis, references: [...references.values()], dynamic, readsText, readsTheme };
}
