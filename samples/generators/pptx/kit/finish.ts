/**
 * What is done to the package after the core saved it: the places where the core stamps the
 * clock or names a font the samples may not use, and the paragraph-build timings it cannot write.
 * Each step is a reported gap of the core's save, kept here so the decks themselves stay clean.
 */
import { FONTS } from '../../lib/fiction';
import { normalizePackage, patchPackage, readPackage } from '../../lib/package';
import { seeded } from '../../lib/random';
import type { Slide } from './slide';
import { corePropertiesXml, type TemplateSpec } from './template';
import { timingXml } from './timing';

const decoder = new TextDecoder();

/** Font names that must not appear anywhere in a sample. */
const FORBIDDEN =
	/typeface="(Calibri|Calibri Light|Arial|Aptos|Segoe UI|Times New Roman|Cambria|Tahoma|Verdana|Symbol|Wingdings|Courier New|Consolas|Georgia)"/u;

/** U+2014, spelled as a code point because the samples may not contain the character. */
const EM_DASH = String.fromCharCode(0x2014);

/** The embedded chart workbook: its own ZIP dates and its default font are normalised. */
async function finishWorkbook(bytes: Uint8Array): Promise<Uint8Array> {
	const styles = (await readPackage(bytes)).get('xl/styles.xml');
	const patched = styles
		? await patchPackage(bytes, {
				'xl/styles.xml': decoder.decode(styles).replaceAll('"Calibri"', `"${FONTS.body}"`),
			})
		: bytes;
	return normalizePackage(patched);
}

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

/**
 * A GUID from a serial number. Every field varies: PowerPoint refuses a diagram whose model ids
 * share their first field, so a constant prefix with a counter at the end does not open.
 */
function stableGuid(serial: number): string {
	const random = seeded(0x5a17 + serial);
	const hex = (digits: number) =>
		Array.from({ length: digits }, () => random.int(0, 15).toString(16).toUpperCase()).join('');
	return `{${hex(8)}-${hex(4)}-4${hex(3)}-A${hex(3)}-${hex(12)}}`;
}

/**
 * The core draws SmartArt model ids and table column ids from the runtime's random source.
 * They are renumbered in order of first appearance so that two builds give the same bytes.
 */
class StableIds {
	private readonly guids = new Map<string, string>();
	private columns = 0;

	diagram(xml: string): string {
		return xml.replace(/(?<!uri=")\{[0-9A-F]{8}-(?:[0-9A-F]{4}-){3}[0-9A-F]{12}\}/giu, (guid) => {
			let stable = this.guids.get(guid);
			if (!stable) {
				stable = stableGuid(this.guids.size + 1);
				this.guids.set(guid, stable);
			}
			return stable;
		});
	}

	slide(xml: string): string {
		return xml.replace(/<a16:(colId|rowId) val="\d+"/gu, (_, kind: string) => {
			this.columns += 1;
			return `<a16:${kind} val="${20260115 + this.columns}"`;
		});
	}
}

export async function finishPackage(
	saved: Uint8Array,
	spec: TemplateSpec,
	slides: readonly Slide[],
): Promise<Uint8Array> {
	const parts = await readPackage(saved);
	const stable = new StableIds();
	const changes: Record<string, string | Uint8Array> = {
		'docProps/core.xml': corePropertiesXml(spec.title, spec.subject),
	};
	for (const [name, bytes] of parts) {
		if (name.startsWith('ppt/embeddings/') && name.endsWith('.xlsx'))
			changes[name] = await finishWorkbook(bytes);
		// The SmartArt drawing the core writes names the Office default font on every run.
		if (/^ppt\/diagrams\/[a-zA-Z]+\d+\.xml$/u.test(name))
			changes[name] = stable.diagram(
				decoder.decode(bytes).replaceAll('typeface="Calibri"', 'typeface="+mn-lt"'),
			);
		if (/^ppt\/slides\/slide\d+\.xml$/u.test(name))
			changes[name] = stable.slide(decoder.decode(bytes));
	}
	for (const slide of slides) {
		if (slide.effects.length === 0) continue;
		const name = `ppt/slides/slide${slide.number}.xml`;
		const xml = changes[name];
		if (typeof xml !== 'string') throw new Error(`The saved package has no ${name}.`);
		changes[name] = withTiming(xml, slide);
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
