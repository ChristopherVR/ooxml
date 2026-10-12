import { readFileSync, writeFileSync } from 'node:fs';

const [input, output] = process.argv.slice(2);
if (!input || !output)
	throw new Error('Pass a native case corpus JSON and an output TypeScript path.');
const corpus = JSON.parse(readFileSync(input, 'utf8').replace(/^\uFEFF/, ''));
let source = `/** Native Visio 16 FIELDPICTURE casing, measured by scripts/visio-field-case-acceptance.ps1.
 * Ranges encode start, end, stride and code-point offset. Unmapped characters stay unchanged.
 * This is format-specific legacy casing, not JavaScript Unicode string casing. */
`;
for (const [code, name] of [
	['38', 'LOWER'],
	['39', 'UPPER'],
]) {
	const mappings = new Map();
	for (const sample of corpus.cases) {
		const characters = [...sample.input],
			converted = [...sample[code]];
		if (characters.length !== converted.length)
			throw new Error('Native casing changed scalar count.');
		characters.forEach((character, index) => {
			const point = character.codePointAt(0),
				next = converted[index].codePointAt(0);
			if (mappings.has(point) && mappings.get(point) !== next)
				throw new Error('Inconsistent native mapping.');
			mappings.set(point, next);
		});
	}
	const pairs = [...mappings].filter(([from, to]) => from !== to).sort((a, b) => a[0] - b[0]);
	source += `const ${name}: readonly (readonly [number, number, number, number])[] = [\n`;
	for (let index = 0; index < pairs.length;) {
		const first = pairs[index];
		let step = pairs[index + 1]?.[0] - first[0];
		if (step !== 1 && step !== 2) step = 1;
		let end = index + 1;
		while (
			end < pairs.length &&
			pairs[end][0] - pairs[end - 1][0] === step &&
			pairs[end][1] - pairs[end][0] === first[1] - first[0]
		)
			end++;
		source += `\t[0x${first[0].toString(16)}, 0x${pairs[end - 1][0].toString(16)}, ${step}, ${first[1] - first[0]}],\n`;
		index = end;
	}
	source += '];\n';
}
source += `
export function visioFieldCase(text: string, upper: boolean): string {
 const ranges = upper ? UPPER : LOWER;
 return Array.from(text, character => {
  const point = character.codePointAt(0)!;
  let low = 0, high = ranges.length;
  while (low < high) {
   const middle = (low + high) >>> 1;
   if (ranges[middle]![1] < point) low = middle + 1;
   else high = middle;
  }
  const range = ranges[low];
  return range && point >= range[0] && (point - range[0]) % range[2] === 0
   ? String.fromCodePoint(point + range[3]) : character;
 }).join('');
}
`;
writeFileSync(output, source);
