// Test-only: a large ground-truth deck cut down to the slides a test reads.
import { readFileSync } from 'node:fs';

import JSZip from 'jszip';

const SLIDE_LIST = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/;

/**
 * The deck at `path` listing only `slideNumbers` (1-based, kept in that order). Loading and saving
 * cost grows with every listed slide, so a test that reads one slide of a hundred-slide deck pays
 * for one. The other slides' parts stay in the package, unlisted and with their part names, so
 * `ppt/diagrams/drawing6.xml` is still the drawing of the deck's slide 6.
 */
export async function deckWithSlides(
	path: string,
	slideNumbers: readonly number[],
): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(readFileSync(path));
	const presentation = await zip.file('ppt/presentation.xml')?.async('string');
	const list = presentation === undefined ? null : SLIDE_LIST.exec(presentation);
	if (!presentation || !list) throw new Error(`${path} lists no slides`);
	const ids = list[1]?.match(/<p:sldId\b[^>]*\/>/g) ?? [];
	const kept = slideNumbers.map((number) => {
		const id = ids[number - 1];
		if (!id) throw new Error(`${path} has no slide ${number}`);
		return id;
	});
	zip.file(
		'ppt/presentation.xml',
		presentation.replace(list[0], `<p:sldIdLst>${kept.join('')}</p:sldIdLst>`),
	);
	return zip.generateAsync({ type: 'arraybuffer', compression: 'STORE' });
}
