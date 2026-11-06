import type { BuiltRun } from './paragraph-run-build';

const ZERO_WIDTH_SPACE = '​';

/**
 * Gives a paragraph that ends in a soft line break (`a:br`) the empty line
 * PowerPoint draws after it.
 *
 * PowerPoint writes `<a:r/>...<a:br/>` for a paragraph whose author pressed
 * Shift+Enter at its end, and shows a blank line of the break's own size under
 * the text (issue #249). A `<br>` that is the last thing in a CSS block adds no
 * line, so those paragraphs lost the line and the text below sat too high.
 * A zero-width space carrying the break's font metrics, after the `<br>`,
 * gives the empty line its height from the break's size, as PowerPoint does,
 * in every binding because it is an ordinary run.
 */
export function appendTrailingBreakLine(runs: BuiltRun[]): void {
	const last = runs[runs.length - 1];
	if (!last || last.text !== '\n') {
		return;
	}
	const style: BuiltRun['style'] = {};
	for (const key of ['fontSize', 'fontFamily', 'fontWeight', 'fontStyle', 'lineHeight'] as const) {
		const value = last.style[key];
		if (value !== undefined) {
			Object.assign(style, { [key]: value });
		}
	}
	// Same segment as the break, so a binding that resolves a run's font from its segment
	// (React does) paints the empty line at the break's size, not the body default.
	runs.push({
		text: ZERO_WIDTH_SPACE,
		style,
		...(last.segmentIndex !== undefined ? { segmentIndex: last.segmentIndex } : {}),
	});
}
