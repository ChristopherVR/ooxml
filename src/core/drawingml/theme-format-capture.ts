import { NS, first, parseXml } from '../xml/index';
import { parseThemeFormatScheme } from './format-scheme';

export function readCapturedFormatScheme(xml: string) {
	const root = parseXml(`<capture xmlns:a="${NS.a}">${xml}</capture>`).documentElement;
	const format = parseThemeFormatScheme(first(root, 'fmtScheme', NS.a));
	if (!format) throw new Error('Invalid captured DrawingML format scheme.');
	return format;
}
