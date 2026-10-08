// Canonical DrawingML text body of a chart (`c:rich`, `c:txPr` children) from the modelled fields:
// `a:bodyPr` with the anchor and insets, an empty `a:lstStyle`, and per paragraph the alignment,
// spacing and default run properties (always written, as Office does in charts) and the runs.
// Text bodies read from a part keep their source instead (`write-shape.ts`).
import { drawingColorXml } from '../drawingml/write-color';
import type {
	DrawingTextBody,
	DrawingTextParagraph,
	DrawingTextRun,
	DrawingTextSpacing,
} from '../drawingml/types';
import { escapeAttribute, escapeText } from './write-util';

type RunProperties = Omit<DrawingTextRun, 'text'>;

function runProperties(local: string, properties: RunProperties | undefined): string {
	const p = properties ?? {};
	const attributes = [
		p.sizePt === undefined ? '' : ` sz="${Math.round(p.sizePt * 100)}"`,
		p.bold === undefined ? '' : ` b="${p.bold ? 1 : 0}"`,
		p.italic === undefined ? '' : ` i="${p.italic ? 1 : 0}"`,
		p.underline === undefined ? '' : ` u="${p.underline ? 'sng' : 'none'}"`,
	].join('');
	const children =
		(p.color ? `<a:solidFill>${drawingColorXml(p.color)}</a:solidFill>` : '') +
		(p.typeface ? `<a:latin typeface="${escapeAttribute(p.typeface)}"/>` : '');
	return children
		? `<a:${local}${attributes}>${children}</a:${local}>`
		: `<a:${local}${attributes}/>`;
}

function spacing(local: string, value: DrawingTextSpacing | undefined): string {
	if (!value) return '';
	const inner =
		value.unit === 'points'
			? `<a:spcPts val="${Math.round(value.value * 100)}"/>`
			: `<a:spcPct val="${Math.round(value.value * 100000)}"/>`;
	return `<a:${local}>${inner}</a:${local}>`;
}

function run(value: DrawingTextRun): string {
	const { text, ...properties } = value;
	const hasProperties = Object.keys(properties).length > 0;
	if (text === '\n')
		return hasProperties ? `<a:br>${runProperties('rPr', properties)}</a:br>` : '<a:br/>';
	return `<a:r>${hasProperties ? runProperties('rPr', properties) : ''}<a:t>${escapeText(text)}</a:t></a:r>`;
}

function paragraph(value: DrawingTextParagraph): string {
	const align = value.align ? ` algn="${escapeAttribute(value.align)}"` : '';
	const properties =
		spacing('lnSpc', value.lineSpacing) +
		spacing('spcBef', value.spaceBefore) +
		spacing('spcAft', value.spaceAfter) +
		runProperties('defRPr', value.defaultProperties);
	return `<a:p><a:pPr${align}>${properties}</a:pPr>${value.runs.map(run).join('')}</a:p>`;
}

/** The children of a chart text body written from the model. */
export function writeTextBody(body: Omit<DrawingTextBody, 'text'> & { text?: string }): string {
	const insets = body.insetsEmu ?? {};
	const attributes = [
		['lIns', insets.left],
		['tIns', insets.top],
		['rIns', insets.right],
		['bIns', insets.bottom],
	]
		.map(([name, value]) => (value === undefined ? '' : ` ${name}="${Math.round(Number(value))}"`))
		.join('');
	const anchor = body.anchor ? ` anchor="${escapeAttribute(body.anchor)}"` : '';
	const paragraphs = body.paragraphs.length
		? body.paragraphs.map(paragraph).join('')
		: paragraph({ runs: [] });
	return `<a:bodyPr${attributes}${anchor}/><a:lstStyle/>${paragraphs}`;
}
