/**
 * The show toolbar's "has annotations" flag counts ink on any slide, as Vue, Svelte, Angular and
 * vanilla do; its Erase button (which clears the slide on screen) follows that slide's ink. React passed the current slide's strokes, so leaving a slide with ink greyed out the
 * toolbar's ink control on the next one while the keep-or-discard prompt still counted it.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, test } from 'vitest';

const source = readFileSync(
	path.join(path.dirname(fileURLToPath(import.meta.url)), 'ViewerCanvasArea.tsx'),
	'utf8',
);

test('passes the any-slide ink flag to the show toolbar', () => {
	expect(source).toContain('hasAnnotations={annotations.hasAnyAnnotations}');
	expect(source).not.toContain('hasAnnotations={annotations.annotationStrokes');
});

test('enables Erase All Ink on Slide from the current slide only', () => {
	expect(source).toContain('hasSlideAnnotations={annotations.annotationStrokes.length > 0}');
});
