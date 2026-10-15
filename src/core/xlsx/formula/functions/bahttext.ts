import { ERR, fail } from '../values';
import { num, spec } from './helpers';

const DIGITS = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
const PLACES = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
const MILLION = 'ล้าน';

/** Words for a number below one million (no zero word). */
function belowMillion(n: number): string {
	const digits = String(n).split('').map(Number);
	const len = digits.length;
	let out = '';
	digits.forEach((d, i) => {
		const place = len - 1 - i;
		if (d === 0) return;
		if (place === 0) out += d === 1 && len > 1 ? 'เอ็ด' : (DIGITS[d] ?? '');
		else if (place === 1)
			out += (d === 1 ? '' : d === 2 ? 'ยี่' : (DIGITS[d] ?? '')) + (PLACES[1] ?? '');
		else out += (DIGITS[d] ?? '') + (PLACES[place] ?? '');
	});
	return out;
}

/** Thai words for a non-negative integer (millions group recursively). */
export function thaiInteger(n: number): string {
	if (n === 0) return DIGITS[0] as string;
	const high = Math.floor(n / 1_000_000);
	const low = n % 1_000_000;
	return (high > 0 ? thaiInteger(high) + MILLION : '') + (low > 0 ? belowMillion(low) : '');
}

export const BAHTTEXT = spec(
	'BAHTTEXT',
	'Text',
	'BAHTTEXT(number)',
	'A number as Thai text with the baht suffix.',
	1,
	1,
	(args) => {
		const value = num(args[0]);
		const negative = value < 0;
		const satangTotal = Math.round(Math.abs(value) * 100);
		if (!Number.isSafeInteger(satangTotal)) fail(ERR.VALUE);
		const baht = Math.floor(satangTotal / 100);
		const satang = satangTotal % 100;
		let text = '';
		if (baht > 0 || satang === 0) text += thaiInteger(baht) + 'บาท';
		text += satang === 0 ? 'ถ้วน' : thaiInteger(satang) + 'สตางค์';
		return (negative && satangTotal > 0 ? 'ลบ' : '') + text;
	},
);
