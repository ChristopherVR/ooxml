// Render generated before/after diagram JSON with the same SVG builder as Word and Excel.
// Run from src/ui: bun scripts/native-smartart.ts <fixture-directory> <capture-directory>
import { JSDOM } from 'jsdom';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { renderDiagramDrawing } from '../src/smartart/smartart-svg';

const input = resolve(process.argv[2]!);
const output = resolve(process.argv[3]!);
const dom = new JSDOM();
try {
	for (const name of readdirSync(input).filter((name) => name.endsWith('-drawing.json'))) {
		const fixture = JSON.parse(readFileSync(resolve(input, name), 'utf8'));
		const { svg, report } = renderDiagramDrawing(
			dom.window.document,
			fixture.drawing,
			fixture.scheme ?? {},
			fixture.fonts ?? {},
			(part) => fixture.pictureUrls[part],
		);
		svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
		svg.setAttribute('width', String(fixture.width));
		svg.setAttribute('height', String(fixture.height));
		const id = name.includes('-before-drawing')
			? name.replace('-drawing.json', '')
			: name.replace('-drawing.json', '-after');
		writeFileSync(resolve(output, `${id}.svg`), svg.outerHTML);
		writeFileSync(resolve(output, `${id}.report.json`), JSON.stringify(report, null, 2));
	}
} finally {
	dom.window.close();
}
