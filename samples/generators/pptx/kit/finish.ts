/**
 * What is done to the package after the core saved it: the paragraph-build timings the core
 * cannot write yet, then a check that no part names a font the samples may not use. The save
 * itself is reproducible (`fixedDate` and a seeded id source, see `deck.ts`).
 */
import { patchPackage, readPackage } from '../../lib/package';
import type { Slide } from './slide';
import { timingXml } from './timing';

const decoder = new TextDecoder();

/** Font names that must not appear anywhere in a sample. */
const FORBIDDEN =
	/typeface="(Calibri|Calibri Light|Arial|Aptos|Segoe UI|Times New Roman|Cambria|Tahoma|Verdana|Symbol|Wingdings|Courier New|Consolas|Georgia)"/u;

/** U+2014, spelled as a code point because the samples may not contain the character. */
const EM_DASH = String.fromCharCode(0x2014);

/** Swaps the slide's `p:timing` (or adds one) for the hand-written paragraph build. */
function withTiming(slideXml: string, slide: Slide): string {
	const ids = new Map<string, string>();
	for (const match of slideXml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/gu))
		ids.set(match[2]!.replaceAll('&amp;', '&'), match[1]!);
	const timing = timingXml(slide.effects, ids);
	const without = slideXml.replace(/<p:timing>.*<\/p:timing>/su, '');
	// p:timing follows p:clrMapOvr and the transition, and precedes p:extLst (CT_Slide order).
	const at = without.lastIndexOf('<p:extLst>');
	const anchor = at >= 0 ? at : without.lastIndexOf('</p:sld>');
	return without.slice(0, anchor) + timing + without.slice(anchor);
}

export async function finishPackage(
	saved: Uint8Array,
	slides: readonly Slide[],
): Promise<Uint8Array> {
	const parts = await readPackage(saved);
	const changes: Record<string, string | Uint8Array> = {};
	for (const slide of slides) {
		if (slide.effects.length === 0) continue;
		const name = `ppt/slides/slide${slide.number}.xml`;
		const bytes = parts.get(name);
		if (!bytes) throw new Error(`The saved package has no ${name}.`);
		changes[name] = withTiming(decoder.decode(bytes), slide);
	}
	const finished = await patchPackage(saved, changes);
	for (const [name, bytes] of await readPackage(finished)) {
		if (!/\.(?:xml|rels)$/u.test(name)) continue;
		const text = decoder.decode(bytes);
		const font = FORBIDDEN.exec(text);
		if (font) throw new Error(`${name} names the font ${font[1]}.`);
		if (text.includes(EM_DASH)) throw new Error(`${name} contains an em-dash.`);
	}
	return finished;
}
