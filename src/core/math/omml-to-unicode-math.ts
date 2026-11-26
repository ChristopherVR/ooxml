/**
 * Office Math to linear format (UnicodeMath): what Word's Linear command shows and what its
 * equation editor lets the user edit in the Unicode input format. The spelling is the one
 * `unicode-math.ts` builds up again, so a structure survives Linear and back to Professional.
 */
import { stripXmlOrderSuffix } from './omml-node';
import { ACCENT_ALIASES } from './latex-omml-symbols';
import type { XmlRecord } from './omml-to-latex-helpers';
import {
	attrVal,
	childNode,
	collectText,
	contentKeys,
	ensureArr,
	hasAttr,
	isOn,
	readRunText,
	soleChild,
} from './omml-to-latex-helpers';
import { BRACKETS, isOperandChar, LINEAR } from './unicode-math-symbols';

function children(node: XmlRecord | undefined): string {
	if (!node || typeof node !== 'object') return '';
	let out = '';
	for (const key of contentKeys(node))
		for (const item of ensureArr(node[key])) out += element(stripXmlOrderSuffix(key), item);
	return out;
}

const arg = (node: XmlRecord, key: string) => children(childNode(node, key));

/** One operand as typed: a run of letters and digits, or one character. */
const simple = (text: string) =>
	Array.from(text).length === 1 || Array.from(text).every((char) => isOperandChar(char));

/** An argument a build-up operator reads whole: bracketed unless it is one operand. */
const operand = (text: string) => (simple(text) ? text : `(${text})`);

/** A base or n-ary body, which keeps typed brackets, so it is grouped invisibly. */
const grouped = (text: string) =>
	simple(text) || isBracketed(text) ? text : `${LINEAR.groupOpen}${text}${LINEAR.groupClose}`;

function isBracketed(text: string): boolean {
	const chars = Array.from(text);
	const open = chars[0];
	if (!open || !(open in BRACKETS)) return false;
	let depth = 0;
	for (let i = 0; i < chars.length; i++) {
		if (chars[i] === open && (i === 0 || open !== BRACKETS[open])) depth++;
		else if (chars[i] === BRACKETS[open]) depth--;
		if (depth === 0 && i < chars.length - 1) return false;
	}
	return chars[chars.length - 1] === BRACKETS[open];
}

function run(node: XmlRecord): string {
	const text = readRunText(node);
	return isOn(childNode(childNode(node, 'm:rPr'), 'm:nor')) ? `"${text}"` : text;
}

function scripts(node: XmlRecord): string {
	const sub = childNode(node, 'm:sub');
	const sup = childNode(node, 'm:sup');
	return `${node['m:sub'] !== undefined ? `_${operand(children(sub))}` : ''}${node['m:sup'] !== undefined ? `^${operand(children(sup))}` : ''}`;
}

function fraction(node: XmlRecord): string {
	const type = attrVal(childNode(childNode(node, 'm:fPr'), 'm:type'));
	const op =
		type === 'noBar' ? LINEAR.stack : type === 'lin' ? LINEAR.linearFraction : LINEAR.fraction;
	return `${operand(arg(node, 'm:num'))}${op}${operand(arg(node, 'm:den'))}`;
}

function radical(node: XmlRecord): string {
	const base = arg(node, 'm:e');
	const degree = isOn(childNode(childNode(node, 'm:radPr'), 'm:degHide')) ? '' : arg(node, 'm:deg');
	if (degree === '3') return `∛${operand(base)}`;
	if (degree === '4') return `∜${operand(base)}`;
	return degree ? `√(${degree}&${base})` : `√${operand(base)}`;
}

function nary(node: XmlRecord): string {
	const pr = childNode(node, 'm:naryPr');
	const chr = attrVal(childNode(pr, 'm:chr')) || '∫';
	const sub = isOn(childNode(pr, 'm:subHide')) ? '' : arg(node, 'm:sub');
	const sup = isOn(childNode(pr, 'm:supHide')) ? '' : arg(node, 'm:sup');
	const limits = `${sub ? `_${operand(sub)}` : ''}${sup ? `^${operand(sup)}` : ''}`;
	return `${chr}${limits}${LINEAR.naryand}${grouped(arg(node, 'm:e'))}`;
}

function delimiter(node: XmlRecord): string {
	const pr = childNode(node, 'm:dPr');
	const beg = childNode(pr, 'm:begChr');
	const end = childNode(pr, 'm:endChr');
	const open = hasAttr(beg) ? attrVal(beg) || LINEAR.invisibleOpen : '(';
	const close = hasAttr(end) ? attrVal(end) || LINEAR.invisibleClose : ')';
	const sep = childNode(pr, 'm:sepChr');
	const separator = hasAttr(sep) ? attrVal(sep) : '│';
	const inner = ensureArr(node['m:e'])
		.map((e) => children(e))
		.join(separator);
	return `${open}${inner}${close}`;
}

function func(node: XmlRecord): string {
	const name = childNode(node, 'm:fName');
	const sole = soleChild(name);
	// The name is typed bare (`sin`, `log_2`, `lim_(n→∞)`), never quoted as normal text.
	const bare = (container: XmlRecord) => {
		const only = soleChild(container);
		return only?.tag === 'm:r' ? readRunText(only.node) : children(container);
	};
	let head = children(name);
	if (sole?.tag === 'm:r') head = readRunText(sole.node);
	else if (sole?.tag === 'm:limLow' || sole?.tag === 'm:limUpp') {
		const op = sole.tag === 'm:limLow' ? '_' : '^';
		head = `${bare(childNode(sole.node, 'm:e'))}${op}${operand(arg(sole.node, 'm:lim'))}`;
	} else if (sole && ['m:sSup', 'm:sSub', 'm:sSubSup'].includes(sole.tag))
		head = `${bare(childNode(sole.node, 'm:e'))}${scripts(sole.node)}`;
	const body = arg(node, 'm:e');
	return `${head}${LINEAR.funcApply}${simple(body) || isBracketed(body) ? body : `${LINEAR.groupOpen}${body}${LINEAR.groupClose}`}`;
}

function accent(node: XmlRecord): string {
	const raw = attrVal(childNode(childNode(node, 'm:accPr'), 'm:chr')) || '̂';
	const chr = ACCENT_ALIASES[raw] ?? raw;
	return `${operand(arg(node, 'm:e'))}${chr}`;
}

function grid(rows: XmlRecord[], cells: (row: XmlRecord) => XmlRecord[]): string {
	return rows.map((row) => cells(row).map(children).join('&')).join('@');
}

function element(tag: string, node: XmlRecord): string {
	switch (tag) {
		case 'm:r':
			return run(node);
		case 'm:f':
			return fraction(node);
		case 'm:rad':
			return radical(node);
		case 'm:sSup':
		case 'm:sSub':
		case 'm:sSubSup':
			return `${grouped(arg(node, 'm:e'))}${scripts(node)}`;
		case 'm:sPre':
			// Scripts with no base apply to the operand after them; the spaces keep them apart.
			return ` ${scripts(node)} ${grouped(arg(node, 'm:e'))}`;
		case 'm:nary':
			return nary(node);
		case 'm:d':
			return delimiter(node);
		case 'm:func':
			return func(node);
		case 'm:acc':
			return accent(node);
		case 'm:bar':
			return `${attrVal(childNode(childNode(node, 'm:barPr'), 'm:pos')) === 'bot' ? LINEAR.underbar : LINEAR.overbar}(${arg(node, 'm:e')})`;
		case 'm:groupChr': {
			const pr = childNode(node, 'm:groupChrPr');
			const chr = attrVal(childNode(pr, 'm:chr')) || '⏟';
			return chr === '⏞' || chr === '⏟' ? `${chr}(${arg(node, 'm:e')})` : arg(node, 'm:e');
		}
		case 'm:limLow':
		case 'm:limUpp':
			return `${operand(arg(node, 'm:e'))}${tag === 'm:limLow' ? LINEAR.below : LINEAR.above}${operand(arg(node, 'm:lim'))}`;
		case 'm:borderBox':
			return `${LINEAR.box}(${arg(node, 'm:e')})`;
		case 'm:m':
			return `${LINEAR.matrix}(${grid(ensureArr(node['m:mr']), (row) => ensureArr(row['m:e']))})`;
		case 'm:eqArr':
			return `${LINEAR.eqArray}(${ensureArr(node['m:e']).map(children).join('@')})`;
		case 'm:box':
		case 'm:phant':
			return ensureArr(node['m:e']).map(children).join('');
		case 'm:oMath':
			return children(node);
		default:
			return collectText(node);
	}
}

/** Linear (UnicodeMath) text for an Office Math tree at `m:oMathPara`, `m:oMath` or below. */
export function convertOmmlToUnicodeMath(omml: Record<string, unknown>): string {
	if (!omml || typeof omml !== 'object') return '';
	const para = omml['m:oMathPara'] as XmlRecord | XmlRecord[] | undefined;
	const roots = para
		? ensureArr(para).flatMap((p) => ensureArr(p['m:oMath']))
		: ensureArr(omml['m:oMath']);
	if (roots.length) return roots.map(children).join(' ');
	return children(omml);
}
