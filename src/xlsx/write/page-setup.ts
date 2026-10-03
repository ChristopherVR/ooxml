import type { PageSetup, Worksheet } from '../model.js';
import {
	HEADER_FOOTER_FLAGS,
	HEADER_FOOTER_TEXT,
	readHeaderFooter,
	readPageSetup,
	type HeaderFooterFields,
} from '../read/sheet-props.js';
import { mergedAttrs } from './attr-merge.js';
import { sameModel, snapshotElement, snapshotXml } from './snapshot.js';
import { el, encodeEscapes, escapeText } from './xml-out.js';

const headerFooterOf = (page: PageSetup): HeaderFooterFields => {
	const out: HeaderFooterFields = {};
	for (const key of HEADER_FOOTER_FLAGS) if (page[key] !== undefined) out[key] = page[key];
	for (const key of Object.keys(HEADER_FOOTER_TEXT) as (keyof typeof HEADER_FOOTER_TEXT)[])
		if (page[key] !== undefined) out[key] = page[key];
	return out;
};

/** `pageSetup`: the source element's attributes kept, the modelled ones patched. */
function setupXml(sheet: Worksheet, page: PageSetup): string {
	const source = snapshotElement(sheet, 'pageSetup');
	const fit = page.fitToWidth !== undefined || page.fitToHeight !== undefined;
	const pick = (p: PageSetup) => ({
		orientation: p.orientation,
		paperSize: p.paperSize,
		scale: p.scale,
		fitToWidth: p.fitToWidth,
		fitToHeight: p.fitToHeight,
	});
	const modelled =
		fit ||
		page.orientation !== undefined ||
		page.paperSize !== undefined ||
		page.scale !== undefined;
	if (!source && !modelled) return '';
	if (source) {
		const before = readPageSetup(undefined, source, undefined, fit) ?? {};
		if (sameModel(pick(before), pick(page))) return snapshotXml(sheet, 'pageSetup') ?? '';
	}
	// `fitToWidth`/`fitToHeight` only mean something with `fitToPage`; without it they are kept.
	const fitValues = fit ? { fitToWidth: page.fitToWidth, fitToHeight: page.fitToHeight } : {};
	return `<pageSetup${mergedAttrs(source, {
		paperSize: page.paperSize,
		scale: page.scale,
		...fitValues,
		orientation: page.orientation,
	})}/>`;
}

/** `headerFooter`: every header and footer and its flags; unmodelled attributes are kept. */
function headerFooterXml(sheet: Worksheet, page: PageSetup): string {
	const source = snapshotElement(sheet, 'headerFooter');
	const fields = headerFooterOf(page);
	if (source && sameModel(readHeaderFooter(source), fields))
		return snapshotXml(sheet, 'headerFooter') ?? '';
	if (!Object.keys(fields).length) return '';
	const children = Object.entries(HEADER_FOOTER_TEXT)
		.map(([key, local]) => {
			const text = fields[key as keyof typeof HEADER_FOOTER_TEXT];
			return text === undefined ? '' : `<${local}>${escapeText(encodeEscapes(text))}</${local}>`;
		})
		.join('');
	const flags: Record<string, boolean | undefined> = {};
	for (const key of HEADER_FOOTER_FLAGS) flags[key] = fields[key];
	const head = mergedAttrs(source, flags);
	return children ? `<headerFooter${head}>${children}</headerFooter>` : `<headerFooter${head}/>`;
}

/** `pageMargins`, `pageSetup` and `headerFooter`, reusing the source XML where unchanged. */
export function pageXml(sheet: Worksheet): {
	margins: string;
	setup: string;
	headerFooter: string;
} {
	const page = sheet.pageSetup ?? {};
	const result = { margins: '', setup: setupXml(sheet, page), headerFooter: '' };
	const marginsSource = snapshotElement(sheet, 'pageMargins');
	if (page.margins) {
		const same =
			marginsSource &&
			sameModel(readPageSetup(marginsSource, undefined, undefined, false)?.margins, page.margins);
		result.margins = same
			? (snapshotXml(sheet, 'pageMargins') ?? '')
			: el('pageMargins', { ...page.margins });
	}
	result.headerFooter = headerFooterXml(sheet, page);
	return result;
}
