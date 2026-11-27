/**
 * The empty presentation a deck starts from: theme, one designed slide master, its layouts, a
 * notes master and the package plumbing. The core's `PresentationBuilder` only offers the stock
 * eleven layouts on a bare master, so this part of the package is written by hand; the core then
 * loads it and adds the slides.
 */
import { FONTS } from '../../lib/fiction';
import { xmlEscape, type PackageParts } from '../../lib/package';
import { SAMPLE_AUTHOR, SAMPLE_DATE_ISO } from '../../lib/types';
import { SLIDE_CX, SLIDE_CY } from './geometry';
import { notesMasterXml } from './notes-master';
import { themeXml, type DeckTheme } from './theme';
import { levelXml, placeholderXml, type LevelStyle, type PlaceholderSpec } from './xml';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS =
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument';
const GROUP =
	'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/** What a master or layout draws besides its placeholders. */
export interface PartArt {
	/** A `p:bg` element; the master must have one, a layout inherits when it has none. */
	background?: string;
	/** Decorative shapes, drawn below the placeholders. */
	shapes?: string;
	/** Pictures the fragments above reference: relationship id to a name in `media`. */
	images?: Record<string, string>;
}

export interface LayoutSpec extends PartArt {
	name: string;
	/** `ST_SlideLayoutType`. */
	type: 'title' | 'obj' | 'secHead' | 'twoObj' | 'titleOnly' | 'blank' | 'objTx' | 'picTx' | 'cust';
	placeholders: PlaceholderSpec[];
	/** Hide the master's decorative shapes on this layout. */
	hideMasterShapes?: boolean;
}

export interface TemplateSpec {
	theme: DeckTheme;
	title: string;
	subject: string;
	master: PartArt & {
		placeholders: PlaceholderSpec[];
		titleStyle: LevelStyle;
		bodyStyle: LevelStyle[];
	};
	layouts: LayoutSpec[];
	/** Picture bytes by file name (`backdrop.png`), stored under `ppt/media/`. */
	media?: Record<string, Uint8Array>;
}

function relationships(entries: [id: string, type: string, target: string][]): string {
	return `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries
		.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`)
		.join('')}</Relationships>`;
}

function imageRelationships(art: PartArt): [string, string, string][] {
	return Object.entries(art.images ?? {}).map(([id, file]) => [
		id,
		`${REL}/image`,
		`../media/${file}`,
	]);
}

function masterXml(spec: TemplateSpec): string {
	const { master, theme } = spec;
	const map = theme.dark
		? 'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"'
		: 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"';
	const layoutIds = spec.layouts
		.map((_, index) => `<p:sldLayoutId id="${2147483649 + index}" r:id="rId${index + 1}"/>`)
		.join('');
	const other = levelXml(1, { size: 14, color: 'tx1' }, FONTS.body);
	return (
		`${XML}<p:sldMaster ${NS}><p:cSld>${master.background ?? ''}<p:spTree>${GROUP}${master.shapes ?? ''}` +
		`${master.placeholders.map((ph) => placeholderXml(ph, FONTS.body)).join('')}</p:spTree></p:cSld>` +
		`<p:clrMap ${map} accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
		`<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst>` +
		`<p:txStyles><p:titleStyle>${levelXml(1, master.titleStyle, FONTS.body)}</p:titleStyle>` +
		`<p:bodyStyle>${master.bodyStyle.map((level, index) => levelXml(index + 1, level, FONTS.body)).join('')}</p:bodyStyle>` +
		`<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${other}</p:otherStyle></p:txStyles></p:sldMaster>`
	);
}

function layoutXml(layout: LayoutSpec): string {
	const type = layout.type === 'cust' ? '' : ` type="${layout.type}"`;
	const hide = layout.hideMasterShapes ? ' showMasterSp="0"' : '';
	return (
		`${XML}<p:sldLayout ${NS}${type} preserve="1" userDrawn="1"${hide}><p:cSld name="${xmlEscape(layout.name)}">${layout.background ?? ''}` +
		`<p:spTree>${GROUP}${layout.shapes ?? ''}${layout.placeholders.map((ph) => placeholderXml(ph, FONTS.body)).join('')}</p:spTree></p:cSld>` +
		'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
	);
}

function presentationXml(): string {
	const levels = Array.from({ length: 9 }, (_, index) =>
		levelXml(index + 1, { size: 18, color: 'tx1', margin: index * 48 }, FONTS.body),
	).join('');
	return (
		`${XML}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
		'<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>' +
		`<p:sldSz cx="${SLIDE_CX}" cy="${SLIDE_CY}"/><p:notesSz cx="6858000" cy="9144000"/>` +
		`<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${levels}</p:defaultTextStyle></p:presentation>`
	);
}

function contentTypes(spec: TemplateSpec): string {
	const layouts = spec.layouts
		.map(
			(_, index) =>
				`<Override PartName="/ppt/slideLayouts/slideLayout${index + 1}.xml" ContentType="${CT}.presentationml.slideLayout+xml"/>`,
		)
		.join('');
	return (
		`${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
		'<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
		`<Override PartName="/ppt/presentation.xml" ContentType="${CT}.presentationml.presentation.main+xml"/>` +
		`<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT}.presentationml.slideMaster+xml"/>${layouts}` +
		`<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${CT}.presentationml.notesMaster+xml"/>` +
		`<Override PartName="/ppt/theme/theme1.xml" ContentType="${CT}.theme+xml"/>` +
		`<Override PartName="/ppt/theme/theme2.xml" ContentType="${CT}.theme+xml"/>` +
		`<Override PartName="/ppt/presProps.xml" ContentType="${CT}.presentationml.presProps+xml"/>` +
		`<Override PartName="/ppt/viewProps.xml" ContentType="${CT}.presentationml.viewProps+xml"/>` +
		`<Override PartName="/ppt/tableStyles.xml" ContentType="${CT}.presentationml.tableStyles+xml"/>` +
		'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
		`<Override PartName="/docProps/app.xml" ContentType="${CT}.extended-properties+xml"/></Types>`
	);
}

/** `docProps/core.xml` of the template, dated the sample date. */
export function corePropertiesXml(title: string, subject: string): string {
	return (
		`${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
		`<dc:title>${xmlEscape(title)}</dc:title><dc:subject>${xmlEscape(subject)}</dc:subject><dc:creator>${SAMPLE_AUTHOR}</dc:creator>` +
		`<cp:lastModifiedBy>${SAMPLE_AUTHOR}</cp:lastModifiedBy><cp:revision>1</cp:revision>` +
		`<dcterms:created xsi:type="dcterms:W3CDTF">${SAMPLE_DATE_ISO}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${SAMPLE_DATE_ISO}</dcterms:modified></cp:coreProperties>`
	);
}

/** The parts of the empty presentation, in package order. */
export function buildTemplate(spec: TemplateSpec): PackageParts {
	const theme = themeXml(spec.theme);
	const parts: PackageParts = {
		'[Content_Types].xml': contentTypes(spec),
		'_rels/.rels': relationships([
			['rId1', `${REL}/officeDocument`, 'ppt/presentation.xml'],
			[
				'rId2',
				'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
				'docProps/core.xml',
			],
			['rId3', `${REL}/extended-properties`, 'docProps/app.xml'],
		]),
		'docProps/core.xml': corePropertiesXml(spec.title, spec.subject),
		'docProps/app.xml': `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ooxml-core</Application><PresentationFormat>Widescreen</PresentationFormat></Properties>`,
		'ppt/presentation.xml': presentationXml(),
		'ppt/_rels/presentation.xml.rels': relationships([
			['rId1', `${REL}/slideMaster`, 'slideMasters/slideMaster1.xml'],
			['rId2', `${REL}/notesMaster`, 'notesMasters/notesMaster1.xml'],
			['rId3', `${REL}/theme`, 'theme/theme1.xml'],
			['rId4', `${REL}/presProps`, 'presProps.xml'],
			['rId5', `${REL}/viewProps`, 'viewProps.xml'],
			['rId6', `${REL}/tableStyles`, 'tableStyles.xml'],
		]),
		'ppt/presProps.xml': `${XML}<p:presentationPr ${NS}/>`,
		'ppt/viewProps.xml': `${XML}<p:viewPr ${NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="80000"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0"><p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:cSldViewPr></p:slideViewPr></p:viewPr>`,
		'ppt/tableStyles.xml': `${XML}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
		'ppt/theme/theme1.xml': theme,
		'ppt/theme/theme2.xml': theme,
		'ppt/slideMasters/slideMaster1.xml': masterXml(spec),
		'ppt/slideMasters/_rels/slideMaster1.xml.rels': relationships([
			...spec.layouts.map((_, index): [string, string, string] => [
				`rId${index + 1}`,
				`${REL}/slideLayout`,
				`../slideLayouts/slideLayout${index + 1}.xml`,
			]),
			[`rId${spec.layouts.length + 1}`, `${REL}/theme`, '../theme/theme1.xml'],
			...imageRelationships(spec.master),
		]),
		'ppt/notesMasters/notesMaster1.xml': notesMasterXml(),
		'ppt/notesMasters/_rels/notesMaster1.xml.rels': relationships([
			['rId1', `${REL}/theme`, '../theme/theme2.xml'],
		]),
	};
	spec.layouts.forEach((layout, index) => {
		parts[`ppt/slideLayouts/slideLayout${index + 1}.xml`] = layoutXml(layout);
		parts[`ppt/slideLayouts/_rels/slideLayout${index + 1}.xml.rels`] = relationships([
			['rId1', `${REL}/slideMaster`, '../slideMasters/slideMaster1.xml'],
			...imageRelationships(layout),
		]);
	});
	for (const [file, bytes] of Object.entries(spec.media ?? {})) parts[`ppt/media/${file}`] = bytes;
	return parts;
}
