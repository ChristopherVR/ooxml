/**
 * What a drawing is made from, and the package parts `Drawing` writes itself because the core has
 * no writer for them: the style sheets, the masters and the theme. Each function here returns the
 * parts to change; `Drawing` applies them.
 */
import { readPackage } from '../../lib/package';
import type { MasterDef, Stencil } from './masters';
import { insertBefore, withContentTypes, withRelationship } from './parts';
import type { StyleBook, StyleDef } from './styles';

export interface PageSpec {
	name: string;
	background?: boolean;
	/** Name of the background page shown behind this page. */
	backPage?: string;
}

export interface DrawingSpec {
	width: number;
	height: number;
	/** Default text and line colour. */
	ink: string;
	styles: readonly StyleDef[];
	masters: readonly MasterDef[];
	pages: readonly PageSpec[];
	title: string;
	subject: string;
	description: string;
	keywords: readonly string[];
	/** The theme part and the PageSheet cells that select it, when the drawing has a theme. */
	theme?: { xml: string; pageCells: string };
}

/** Reads one part of the package being built as text. */
export type ReadPart = (part: string) => Promise<string>;

const decoder = new TextDecoder();
/** Written as a code point so that this file holds none itself. */
const EM_DASH = String.fromCodePoint(0x2014);
/** Fonts a sample must never name (the brief's rule 5). */
const FORBIDDEN_FONTS =
	/Calibri|Aptos|Segoe|Arial|Times New Roman|Cambria|Consolas|Georgia|Verdana|Tahoma|Courier New|Symbol|Wingdings/u;

/** The part of every page, by page id, from the pages part and its relationships. */
export function pagePartNames(pagesXml: string, rels: string): Map<string, string> {
	const parts = new Map<string, string>();
	for (const match of pagesXml.matchAll(/<Page\b[^>]*\bID="(\d+)"[\s\S]*?<\/Page>/gu)) {
		const relId = /<Rel\b[^>]*r:id="([^"]+)"/u.exec(match[0])?.[1];
		const target = new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`, 'u').exec(rels)?.[1];
		const reversed = new RegExp(`Target="([^"]+)"[^>]*Id="${relId}"`, 'u').exec(rels)?.[1];
		const part = target ?? reversed;
		if (!match[1] || !part) throw new Error('A page has no part.');
		parts.set(match[1], `visio/pages/${part}`);
	}
	return parts;
}

/** A page part with `shapes` and `connects` added to what it already holds. */
export function pageWithShapes(page: string, shapes: string, connects: string): string {
	let xml = page.replace(/<PageContents\b([^>]*)\/>/u, '<PageContents$1></PageContents>');
	xml = xml.replace(/<Shapes\s*\/>/u, '<Shapes></Shapes>');
	if (!xml.includes('</Shapes>'))
		xml = xml.replace(/(<PageContents\b[^>]*>)/u, '$1<Shapes></Shapes>');
	xml = insertBefore(xml, '</Shapes>', shapes);
	if (connects) {
		xml = xml.replace(/<Connects\s*\/>/u, '<Connects></Connects>');
		xml = xml.includes('</Connects>')
			? insertBefore(xml, '</Connects>', connects)
			: insertBefore(xml, '</PageContents>', `<Connects>${connects}</Connects>`);
	}
	return xml;
}

/** The document part, the masters and the theme, with their relationships and content types. */
export async function documentPartChanges(
	read: ReadPart,
	stencil: Stencil,
	styles: StyleBook,
	spec: DrawingSpec,
): Promise<Record<string, string>> {
	const changes: Record<string, string> = {
		'visio/document.xml': styles.documentXml(),
		// A master resolves theme formulas against its own page sheet, so it selects the theme too.
		'visio/masters/masters.xml': stencil.mastersXml(spec.theme?.pageCells ?? ''),
		'visio/masters/_rels/masters.xml.rels': stencil.relsXml(),
	};
	const types: [string, string][] = [
		['/visio/masters/masters.xml', 'application/vnd.ms-visio.masters+xml'],
	];
	for (const master of stencil.masters) {
		changes[`visio/masters/master${master.id}.xml`] = stencil.masterXml(master, styles);
		types.push([`/visio/masters/master${master.id}.xml`, 'application/vnd.ms-visio.master+xml']);
	}
	let documentRels = withRelationship(
		await read('visio/_rels/document.xml.rels'),
		'rIdMasters',
		'http://schemas.microsoft.com/visio/2010/relationships/masters',
		'masters/masters.xml',
	);
	if (spec.theme) {
		changes['visio/theme/theme1.xml'] = spec.theme.xml;
		types.push([
			'/visio/theme/theme1.xml',
			'application/vnd.openxmlformats-officedocument.theme+xml',
		]);
		documentRels = withRelationship(
			documentRels,
			'rIdTheme',
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
			'theme/theme1.xml',
		);
		const pagesXml = await read('visio/pages/pages.xml');
		changes['visio/pages/pages.xml'] = pagesXml.replaceAll(
			/(<PageSheet\b[^>]*>)/gu,
			`$1${spec.theme.pageCells}`,
		);
	}
	changes['visio/_rels/document.xml.rels'] = documentRels;
	changes['[Content_Types].xml'] = withContentTypes(await read('[Content_Types].xml'), types);
	return changes;
}

/** Fails when a part names a font no sample may use, or contains an em-dash. */
export async function assertSampleRules(bytes: Uint8Array): Promise<void> {
	for (const [name, data] of await readPackage(bytes)) {
		const text = decoder.decode(data);
		const font = FORBIDDEN_FONTS.exec(text);
		if (font) throw new Error(`${name} names the font ${font[0]}.`);
		if (text.includes(EM_DASH)) throw new Error(`${name} contains an em-dash.`);
	}
}
