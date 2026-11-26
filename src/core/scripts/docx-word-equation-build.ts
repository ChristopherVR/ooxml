// Builds the documents `docx-word-equation-acceptance.ps1` opens in Word. Run with bun from
// src/core:   bun scripts/docx-word-equation-build.ts <output folder>
// One document per Equation tab gallery (every structure, display and inline, with its empty
// placeholders), the built-in equations, Linear-format equations and Unicode linear input built
// up. `manifest.json` lists, per file, how many equations Word must find (`Document.OMaths`).
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createDocument, saveDocx, type Paragraph } from '../docx/index';
import {
	buildUpEquation,
	EQUATION_GALLERY,
	EQUATION_STRUCTURES,
	editableEquation,
	linearEquationXml,
	writeEditedEquation,
} from '../math/equation-tools';
import wordEquations from '../math/__fixtures__/word-equations.json';

const out = process.argv[2];
if (!out) throw new Error('usage: bun scripts/docx-word-equation-build.ts <output folder>');
mkdirSync(out, { recursive: true });
const manifest: Record<string, number> = {};

async function write(
	name: string,
	equations: Array<{ omml: string; display: boolean }>,
	expected = equations.length,
) {
	const model = createDocument();
	model.blocks = equations.map((equation, index): Paragraph => ({
		type: 'paragraph',
		id: `p${index}`,
		runs: equation.display
			? [{ text: '', equation }]
			: [{ text: 'Inline ' }, { text: '', equation }, { text: ' text.' }],
	}));
	writeFileSync(join(out, name), await saveDocx(model));
	manifest[name] = expected;
}

const built = (latex: string, display: boolean, format: 'latex' | 'unicode' = 'latex') => {
	const omml = buildUpEquation(latex, format, display);
	if (!omml) throw new Error(`does not build: ${latex}`);
	return { omml, display };
};

for (const gallery of EQUATION_STRUCTURES)
	await write(
		`structures-${gallery.id}.docx`,
		gallery.structures.flatMap((entry) => [built(entry.latex, true), built(entry.latex, false)]),
	);
await write(
	'built-in.docx',
	EQUATION_GALLERY.map((entry) => built(entry.latex, true)),
);
await write('linear.docx', [
	{ omml: linearEquationXml('a/b+√x', true), display: true },
	{ omml: linearEquationXml('x^2', false), display: false },
]);
await write(
	'unicode-input.docx',
	['(a+b)/c', '∑_(i=1)^n▒a_i', 'lim_(n→∞) a_n', '■(a&b@c&d)', '"if "x≥0', 'e^-x', '√(3&x)'].map(
		(text) => built(text, true, 'unicode'),
	),
);
// Equations Word wrote, edited as linear text: their run fonts and paragraph properties are kept.
const NS =
	' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const edited = wordEquations.flatMap(({ omml }) => {
	const source = omml.replace(/^<(m:oMath(?:Para)?)/, `<$1${NS}`);
	const edit = editableEquation(source, 'unicode');
	const out =
		edit &&
		writeEditedEquation({ ...edit, text: `${edit.text}+1`, format: 'unicode', display: true });
	return out ? [{ omml: out, display: true }] : [];
});
await write('edited-word.docx', edited);
// Word opens an empty equation without repair and drops it while loading.
await write('empty.docx', [built('', true), built('', false)], 0);
writeFileSync(join(out, 'manifest.json'), JSON.stringify(manifest, null, '\t'));
console.log(`wrote ${Object.keys(manifest).length} documents to ${out}`);
