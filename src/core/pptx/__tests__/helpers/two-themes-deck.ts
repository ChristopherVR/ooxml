/**
 * A synthetic deck whose two slide masters use two different theme parts, for the "theme by
 * path" switch (Presentation properties > Theme: pick a theme part already in the package).
 *
 * - `ppt/theme/theme1.xml` ("Tide": accent 1 blue, headings in Georgia) is used by slide master 1,
 *   whose layout slides 1 and 2 use. `ppt/theme/theme2.xml` ("Ember": accent 1 red, headings in
 *   Courier New) is used by slide master 2, whose layout slide 3 uses.
 * - Every slide has a rectangle filled with `accent1` (named {@link TWO_THEMES.shapeName}), and
 *   every master has a thin `accent1` bar ({@link TWO_THEMES.masterBarName}) the slides inherit.
 * - Every scheme slot differs between the two themes except `dk1` and `lt1`, and no slot repeats a
 *   value within a theme, so a colour re-resolved from one theme to the other is unambiguous.
 * - With `scriptFaces`, each part's heading font also names a per-script face the other part lacks
 *   ({@link TWO_THEMES.scriptFaces}), so the deck's theme (which merges every part's script faces)
 *   matches neither part's own font scheme.
 *
 * One definition for the core tests, the binding tests and the browser fixture
 * (`e2e/pptx/fixtures/generate-two-themes-fixture.ts`), so the three cannot drift apart.
 */
import JSZip from 'jszip';

interface DeckTheme {
	readonly path: string;
	readonly name: string;
	readonly colors: Readonly<Record<(typeof SCHEME_KEYS)[number], string>>;
	readonly major: string;
	readonly minor: string;
}

const SCHEME_KEYS = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

const TIDE: DeckTheme = {
	path: 'ppt/theme/theme1.xml',
	name: 'Tide',
	colors: {
		dk1: '000000',
		lt1: 'FFFFFF',
		dk2: '17324D',
		lt2: 'E4EEF5',
		accent1: '1F6FB2',
		accent2: '2A9D8F',
		accent3: '8AB17D',
		accent4: 'E9C46A',
		accent5: '6C5B7B',
		accent6: '355070',
		hlink: '0B5394',
		folHlink: '674EA7',
	},
	major: 'Georgia',
	minor: 'Calibri',
};

const EMBER: DeckTheme = {
	path: 'ppt/theme/theme2.xml',
	name: 'Ember',
	colors: {
		dk1: '000000',
		lt1: 'FFFFFF',
		dk2: '4A1C14',
		lt2: 'FBEDE7',
		accent1: 'C0392B',
		accent2: 'E67E22',
		accent3: 'F1C40F',
		accent4: '7D3C98',
		accent5: 'A04000',
		accent6: '641E16',
		hlink: 'B03A2E',
		folHlink: '884EA0',
	},
	major: 'Courier New',
	minor: 'Arial',
};

/** What the deck holds, for assertions. */
export const TWO_THEMES = {
	theme1: TIDE,
	theme2: EMBER,
	master1: 'ppt/slideMasters/slideMaster1.xml',
	master2: 'ppt/slideMasters/slideMaster2.xml',
	layout1: 'ppt/slideLayouts/slideLayout1.xml',
	layout2: 'ppt/slideLayouts/slideLayout2.xml',
	/** The `accent1` rectangle on every slide. */
	shapeName: 'AccentShape',
	/** The `accent1` bar on every master, inherited by its slides. */
	masterBarName: 'MasterBar',
	/** The per-script heading faces each part names with `scriptFaces`, by theme part. */
	scriptFaces: {
		[TIDE.path]: { Jpan: 'Yu Mincho' },
		[EMBER.path]: { Hang: 'Batang' },
	} as Readonly<Record<string, Readonly<Record<string, string>>>>,
	/** Slide index (0-based) to layout: slides 1 and 2 on master 1, slide 3 on master 2. */
	slideLayouts: [
		'ppt/slideLayouts/slideLayout1.xml',
		'ppt/slideLayouts/slideLayout1.xml',
		'ppt/slideLayouts/slideLayout2.xml',
	],
} as const;

const EMU_PER_PX = 9525;
const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
	'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT_NS = 'application/vnd.openxmlformats-officedocument';

function rels(...entries: readonly (readonly [string, string, string])[]): string {
	const body = entries
		.map(
			([id, type, target]) =>
				`<Relationship Id="${id}" Type="${REL_NS}/${type}" Target="${target}"/>`,
		)
		.join('');
	return `${PROLOG}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

function themeXml(theme: DeckTheme, scriptFaces: boolean): string {
	const scripts = Object.entries(scriptFaces ? (TWO_THEMES.scriptFaces[theme.path] ?? {}) : {})
		.map(([script, typeface]) => `<a:font script="${script}" typeface="${typeface}"/>`)
		.join('');
	const slots = SCHEME_KEYS.map(
		(key) => `<a:${key}><a:srgbClr val="${theme.colors[key]}"/></a:${key}>`,
	);
	const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	const line = (w: number) => `<a:ln w="${w}">${fill}</a:ln>`;
	const effect = '<a:effectStyle><a:effectLst/></a:effectStyle>';
	return (
		`${PROLOG}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${theme.name}">` +
		`<a:themeElements><a:clrScheme name="${theme.name}">${slots.join('')}</a:clrScheme>` +
		`<a:fontScheme name="${theme.name}">` +
		`<a:majorFont><a:latin typeface="${theme.major}"/><a:ea typeface=""/><a:cs typeface=""/>${scripts}</a:majorFont>` +
		`<a:minorFont><a:latin typeface="${theme.minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
		`</a:fontScheme><a:fmtScheme name="${theme.name}">` +
		`<a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst>` +
		`<a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>` +
		`<a:effectStyleLst>${effect}${effect}${effect}</a:effectStyleLst>` +
		`<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst>` +
		`</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
	);
}

function rect(id: number, name: string, x: number, y: number, w: number, h: number): string {
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
		`<p:spPr><a:xfrm><a:off x="${x * EMU_PER_PX}" y="${y * EMU_PER_PX}"/>` +
		`<a:ext cx="${w * EMU_PER_PX}" cy="${h * EMU_PER_PX}"/></a:xfrm>` +
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
		`<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`
	);
}

function heading(id: number, text: string): string {
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Heading"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
		`<p:spPr><a:xfrm><a:off x="${80 * EMU_PER_PX}" y="${60 * EMU_PER_PX}"/>` +
		`<a:ext cx="${1120 * EMU_PER_PX}" cy="${100 * EMU_PER_PX}"/></a:xfrm>` +
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
		`<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="4000">` +
		`<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr>` +
		`<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
	);
}

const TREE_ROOT =
	'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const CLR_MAP =
	'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" ' +
	'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

function masterXml(name: string, layoutId: number): string {
	return (
		`${PROLOG}<p:sldMaster ${NS}><p:cSld name="${name}">` +
		`<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
		`<p:spTree>${TREE_ROOT}${rect(2, TWO_THEMES.masterBarName, 0, 700, 1280, 20)}</p:spTree></p:cSld>` +
		`${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="${layoutId}" r:id="rId1"/></p:sldLayoutIdLst>` +
		`</p:sldMaster>`
	);
}

function layoutXml(name: string): string {
	return (
		`${PROLOG}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="${name}">` +
		`<p:spTree>${TREE_ROOT}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
		`</p:sldLayout>`
	);
}

function slideXml(title: string): string {
	return (
		`${PROLOG}<p:sld ${NS}><p:cSld><p:spTree>${TREE_ROOT}${heading(2, title)}` +
		`${rect(3, TWO_THEMES.shapeName, 400, 250, 480, 270)}</p:spTree></p:cSld>` +
		`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
	);
}

export interface TwoThemesDeckOptions {
	/** Give each part's heading font a per-script face the other part lacks. */
	readonly scriptFaces?: boolean;
}

/** Every part of the deck, by path. */
export function twoThemesDeckParts({ scriptFaces = false }: TwoThemesDeckOptions = {}): Record<
	string,
	string
> {
	const override = (part: string, type: string) =>
		`<Override PartName="/${part}" ContentType="${CT_NS}.${type}+xml"/>`;
	const slides = TWO_THEMES.slideLayouts.map((_, index) => `ppt/slides/slide${index + 1}.xml`);
	const parts: Record<string, string> = {
		'[Content_Types].xml':
			`${PROLOG}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			override('ppt/presentation.xml', 'presentationml.presentation.main') +
			override(TWO_THEMES.master1, 'presentationml.slideMaster') +
			override(TWO_THEMES.master2, 'presentationml.slideMaster') +
			override(TWO_THEMES.layout1, 'presentationml.slideLayout') +
			override(TWO_THEMES.layout2, 'presentationml.slideLayout') +
			override(TIDE.path, 'theme') +
			override(EMBER.path, 'theme') +
			slides.map((slide) => override(slide, 'presentationml.slide')).join('') +
			`</Types>`,
		'_rels/.rels': rels(['rId1', 'officeDocument', 'ppt/presentation.xml']),
		'ppt/presentation.xml':
			`${PROLOG}<p:presentation ${NS}><p:sldMasterIdLst>` +
			`<p:sldMasterId id="2147483648" r:id="rId1"/><p:sldMasterId id="2147483650" r:id="rId2"/>` +
			`</p:sldMasterIdLst><p:sldIdLst>` +
			slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${3 + index}"/>`).join('') +
			`</p:sldIdLst><p:sldSz cx="${1280 * EMU_PER_PX}" cy="${720 * EMU_PER_PX}"/>` +
			`<p:notesSz cx="${720 * EMU_PER_PX}" cy="${1280 * EMU_PER_PX}"/></p:presentation>`,
		'ppt/_rels/presentation.xml.rels': rels(
			['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'],
			['rId2', 'slideMaster', 'slideMasters/slideMaster2.xml'],
			...slides.map(
				(_, index) => [`rId${3 + index}`, 'slide', `slides/slide${index + 1}.xml`] as const,
			),
			[`rId${3 + slides.length}`, 'theme', 'theme/theme1.xml'],
		),
		[TIDE.path]: themeXml(TIDE, scriptFaces),
		[EMBER.path]: themeXml(EMBER, scriptFaces),
		[TWO_THEMES.master1]: masterXml('Tide Master', 2147483649),
		[TWO_THEMES.master2]: masterXml('Ember Master', 2147483651),
		'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels(
			['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'],
			['rId2', 'theme', '../theme/theme1.xml'],
		),
		'ppt/slideMasters/_rels/slideMaster2.xml.rels': rels(
			['rId1', 'slideLayout', '../slideLayouts/slideLayout2.xml'],
			['rId2', 'theme', '../theme/theme2.xml'],
		),
		[TWO_THEMES.layout1]: layoutXml('Tide Blank'),
		[TWO_THEMES.layout2]: layoutXml('Ember Blank'),
		'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([
			'rId1',
			'slideMaster',
			'../slideMasters/slideMaster1.xml',
		]),
		'ppt/slideLayouts/_rels/slideLayout2.xml.rels': rels([
			'rId1',
			'slideMaster',
			'../slideMasters/slideMaster2.xml',
		]),
	};
	TWO_THEMES.slideLayouts.forEach((layout, index) => {
		parts[slides[index]!] = slideXml(`Slide ${index + 1}`);
		parts[`ppt/slides/_rels/slide${index + 1}.xml.rels`] = rels([
			'rId1',
			'slideLayout',
			`../slideLayouts/${layout.slice(layout.lastIndexOf('/') + 1)}`,
		]);
	});
	return parts;
}

/** The deck as `.pptx` bytes. */
export async function buildTwoThemesDeck(options: TwoThemesDeckOptions = {}): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [path, xml] of Object.entries(twoThemesDeckParts(options))) {
		zip.file(path, xml);
	}
	return zip.generateAsync({ type: 'uint8array' });
}

/**
 * Each theme part's heading face (the major latin typeface) in saved `.pptx` bytes, by part path:
 * what a reload draws the headings of that part's masters with.
 */
export async function themePartHeadingFaces(
	bytes: Uint8Array,
): Promise<Record<string, string | undefined>> {
	const zip = await JSZip.loadAsync(bytes);
	const faces: Record<string, string | undefined> = {};
	for (const { path } of [TIDE, EMBER]) {
		const xml = (await zip.file(path)?.async('string')) ?? '';
		faces[path] = /<a:majorFont>\s*<a:latin typeface="([^"]*)"/u.exec(xml)?.[1];
	}
	return faces;
}
