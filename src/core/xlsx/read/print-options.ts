import type { XmlElement } from '../../xml/index.js';
import type { PrintOptions } from '../model.js';
import { boolAttr } from './xml-util.js';

const KEYS = ['gridLines', 'headings', 'horizontalCentered', 'verticalCentered'] as const;

/** Reads `<printOptions>`; undefined when absent or when every option is off. */
export function readPrintOptions(node: XmlElement | undefined): PrintOptions | undefined {
	if (!node) return undefined;
	const out: PrintOptions = {};
	for (const key of KEYS) if (boolAttr(node, key)) out[key] = true;
	return Object.keys(out).length ? out : undefined;
}

/** Writes `<printOptions>` from the model ('' when nothing is set). */
export function printOptionsXml(options: PrintOptions | undefined): string {
	if (!options) return '';
	const attrs = KEYS.filter((key) => options[key]).map((key) => ` ${key}="1"`);
	return attrs.length ? `<printOptions${attrs.join('')}/>` : '';
}
