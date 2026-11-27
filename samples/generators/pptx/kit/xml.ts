/**
 * Hand-written DrawingML fragments for the slide master and the layouts. The pptx SDK builds
 * slides, but it has no designed master or layout of its own, so those parts are authored here.
 */
import { xmlEscape } from '../../lib/package';
import { emu } from './geometry';

/** A hex colour (`#RRGGBB`) or a scheme colour name (`tx1`, `accent2`, `bg2`). */
export type Paint = string;

export function colorXml(paint: Paint, alpha?: number): string {
	const inner = alpha === undefined ? '' : `<a:alpha val="${Math.round(alpha * 100000)}"/>`;
	return paint.startsWith('#')
		? `<a:srgbClr val="${paint.slice(1).toUpperCase()}">${inner}</a:srgbClr>`
		: `<a:schemeClr val="${paint}">${inner}</a:schemeClr>`;
}

export interface GradientStop {
	at: number;
	color: Paint;
	alpha?: number;
}

export type RawFill =
	| { solid: Paint; alpha?: number }
	| { stops: GradientStop[]; angle?: number; radial?: { cx: number; cy: number } }
	| 'none';

export function fillXml(fill: RawFill): string {
	if (fill === 'none') return '<a:noFill/>';
	if ('solid' in fill) return `<a:solidFill>${colorXml(fill.solid, fill.alpha)}</a:solidFill>`;
	const stops = fill.stops
		.map(
			(stop) =>
				`<a:gs pos="${Math.round(stop.at * 100000)}">${colorXml(stop.color, stop.alpha)}</a:gs>`,
		)
		.join('');
	if (fill.radial) {
		const l = Math.round(fill.radial.cx * 100000);
		const t = Math.round(fill.radial.cy * 100000);
		return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:path path="circle"><a:fillToRect l="${l}" t="${t}" r="${100000 - l}" b="${100000 - t}"/></a:path></a:gradFill>`;
	}
	return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round((fill.angle ?? 90) * 60000)}" scaled="0"/></a:gradFill>`;
}

interface Frame {
	x: number;
	y: number;
	w: number;
	h: number;
}

export function xfrm(frame: Frame, rotation = 0): string {
	const rot = rotation ? ` rot="${Math.round(rotation * 60000)}"` : '';
	return `<a:xfrm${rot}><a:off x="${emu(frame.x)}" y="${emu(frame.y)}"/><a:ext cx="${emu(frame.w)}" cy="${emu(frame.h)}"/></a:xfrm>`;
}

export interface RawRun {
	text: string;
	/** Size in points. */
	size: number;
	color: Paint;
	alpha?: number;
	font?: 'major' | 'minor';
	bold?: boolean;
	/** Letter spacing in hundredths of a point. */
	tracking?: number;
	caps?: boolean;
}

export function runXml(run: RawRun): string {
	const face = run.font === 'major' ? 'mj' : 'mn';
	const attrs = [
		'lang="en-US"',
		`sz="${Math.round(run.size * 100)}"`,
		`b="${run.bold ? 1 : 0}"`,
		run.tracking === undefined ? '' : `spc="${run.tracking}"`,
		run.caps ? 'cap="all"' : '',
	]
		.filter(Boolean)
		.join(' ');
	return `<a:r><a:rPr ${attrs}><a:solidFill>${colorXml(run.color, run.alpha)}</a:solidFill><a:latin typeface="+${face}-lt"/></a:rPr><a:t>${xmlEscape(run.text)}</a:t></a:r>`;
}

export interface RawShape extends Frame {
	id: number;
	name: string;
	geometry?: string;
	/** Adjust values (`adj` to formula value). */
	adjust?: Record<string, number>;
	fill?: RawFill;
	line?: { color: Paint; width: number; alpha?: number };
	rotation?: number;
	runs?: RawRun[];
	align?: 'l' | 'ctr' | 'r';
	anchor?: 't' | 'ctr' | 'b';
}

/** A decorative shape or a static text box (never a placeholder). */
export function rawShape(shape: RawShape): string {
	const adjust = Object.entries(shape.adjust ?? {})
		.map(([name, value]) => `<a:gd name="${name}" fmla="val ${value}"/>`)
		.join('');
	const line = shape.line
		? `<a:ln w="${emu(shape.line.width)}"><a:solidFill>${colorXml(shape.line.color, shape.line.alpha)}</a:solidFill></a:ln>`
		: '<a:ln><a:noFill/></a:ln>';
	const text = shape.runs
		? `<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${shape.anchor ?? 't'}"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${shape.align ?? 'l'}"/>${shape.runs.map(runXml).join('')}</a:p></p:txBody>`
		: '';
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${shape.id}" name="${xmlEscape(shape.name)}"/><p:cNvSpPr${shape.runs ? ' txBox="1"' : ''}/><p:nvPr userDrawn="1"/></p:nvSpPr>` +
		`<p:spPr>${xfrm(shape, shape.rotation)}<a:prstGeom prst="${shape.geometry ?? 'rect'}"><a:avLst>${adjust}</a:avLst></a:prstGeom>${fillXml(shape.fill ?? 'none')}${line}</p:spPr>${text}</p:sp>`
	);
}

/** A picture on the master or a layout; `relId` is declared in the part's `images`. */
export function rawPicture(picture: Frame & { id: number; name: string; relId: string }): string {
	return (
		`<p:pic><p:nvPicPr><p:cNvPr id="${picture.id}" name="${xmlEscape(picture.name)}" descr=""/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>` +
		`<p:blipFill><a:blip r:embed="${picture.relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
		`<p:spPr>${xfrm(picture)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
	);
}

/** A slide background (`p:bg`) from a fill. */
export function backgroundXml(fill: RawFill): string {
	return `<p:bg><p:bgPr>${fillXml(fill)}<a:effectLst/></p:bgPr></p:bg>`;
}

/** A background that stretches a picture over the slide. */
export function pictureBackgroundXml(relId: string): string {
	return `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${relId}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
}

/** One level of a text style (`a:lvlNpPr`), sizes in points. */
export interface LevelStyle {
	size: number;
	color: Paint;
	font?: 'major' | 'minor';
	bold?: boolean;
	align?: 'l' | 'ctr' | 'r';
	/** A Unicode bullet drawn in the body font, or none. */
	bullet?: string;
	bulletColor?: Paint;
	/** Left margin and hanging indent in px. */
	margin?: number;
	indent?: number;
	/** Space before, in points. */
	before?: number;
	/** Line spacing in percent. */
	leading?: number;
	tracking?: number;
	caps?: boolean;
}

export function levelXml(level: number, style: LevelStyle, bodyFont: string): string {
	const face = style.font === 'major' ? 'mj' : 'mn';
	const bullet = style.bullet
		? `<a:buClr>${colorXml(style.bulletColor ?? style.color)}</a:buClr><a:buFont typeface="${bodyFont}"/><a:buChar char="${xmlEscape(style.bullet)}"/>`
		: '<a:buNone/>';
	return (
		`<a:lvl${level}pPr marL="${emu(style.margin ?? 0)}" indent="${emu(style.indent ?? 0)}" algn="${style.align ?? 'l'}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
		`<a:lnSpc><a:spcPct val="${Math.round((style.leading ?? 100) * 1000)}"/></a:lnSpc>` +
		`<a:spcBef><a:spcPts val="${Math.round((style.before ?? 0) * 100)}"/></a:spcBef>${bullet}` +
		`<a:defRPr sz="${Math.round(style.size * 100)}" b="${style.bold ? 1 : 0}" kern="1200"${style.tracking === undefined ? '' : ` spc="${style.tracking}"`}${style.caps ? ' cap="all"' : ''}>` +
		`<a:solidFill>${colorXml(style.color)}</a:solidFill><a:latin typeface="+${face}-lt"/><a:ea typeface="+${face}-ea"/><a:cs typeface="+${face}-cs"/></a:defRPr>` +
		`</a:lvl${level}pPr>`
	);
}

export interface PlaceholderSpec extends Frame {
	id: number;
	name: string;
	type: 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'dt' | 'ftr' | 'sldNum';
	idx?: number;
	anchor?: 't' | 'ctr' | 'b';
	/** Per-level overrides of the master style. */
	levels?: LevelStyle[];
	/** The prompt PowerPoint shows in the empty placeholder. */
	prompt?: string;
	/** Shrink text on overflow (`a:normAutofit`) instead of the default of no autofit. */
	shrink?: boolean;
}

/** The text a slide-number field shows until PowerPoint fills it in. */
const NUMBER_FIELD =
	'<a:fld id="{B1E0F6E2-3A51-4C0B-9D2C-5A1E7F3C9A10}" type="slidenum"><a:rPr lang="en-US"/><a:t>&#8249;#&#8250;</a:t></a:fld>';

export function placeholderXml(spec: PlaceholderSpec, bodyFont: string): string {
	const idx = spec.idx === undefined ? '' : ` idx="${spec.idx}"`;
	const small = spec.type === 'dt' || spec.type === 'ftr' || spec.type === 'sldNum';
	const levels = (spec.levels ?? []).map((level, at) => levelXml(at + 1, level, bodyFont)).join('');
	const body =
		spec.type === 'sldNum'
			? NUMBER_FIELD
			: `<a:r><a:rPr lang="en-US"/><a:t>${xmlEscape(spec.prompt ?? '')}</a:t></a:r>`;
	return (
		`<p:sp><p:nvSpPr><p:cNvPr id="${spec.id}" name="${xmlEscape(spec.name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="${spec.type}"${small ? ' sz="quarter"' : ''}${idx}/></p:nvPr></p:nvSpPr>` +
		`<p:spPr>${xfrm(spec)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
		`<p:txBody><a:bodyPr vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="${spec.anchor ?? 't'}">${spec.shrink ? '<a:normAutofit/>' : '<a:noAutofit/>'}</a:bodyPr>` +
		`<a:lstStyle>${levels}</a:lstStyle><a:p>${body}</a:p></p:txBody></p:sp>`
	);
}
