import type { PlaceholderDefinition } from './layout-operations';

/** Starter frames in the deck's own dimensions, expressed in CSS pixels. */
export function standardLayoutPlaceholders(
	type: string,
	width: number,
	height: number,
): PlaceholderDefinition[] {
	if (type === 'blank') return [];
	const box = (
		type: string,
		idx: number,
		x: number,
		y: number,
		w: number,
		h: number,
	): PlaceholderDefinition => ({
		type,
		idx,
		x: x * width,
		y: y * height,
		width: w * width,
		height: h * height,
	});
	const title = box('title', 0, 0.05, 0.04, 0.9, 0.15);
	const content = box('obj', 1, 0.05, 0.23, 0.9, 0.64);
	let slots: PlaceholderDefinition[];
	switch (type) {
		case 'title':
			slots = [
				box('ctrTitle', 0, 0.08, 0.24, 0.84, 0.22),
				box('subTitle', 1, 0.12, 0.49, 0.76, 0.2),
			];
			break;
		case 'secHead':
			slots = [box('title', 0, 0.05, 0.37, 0.9, 0.2), box('body', 1, 0.05, 0.6, 0.9, 0.23)];
			break;
		case 'twoObj':
		case 'twoTxTwoObj':
			slots = [title, box('obj', 1, 0.05, 0.27, 0.43, 0.6), box('obj', 2, 0.52, 0.27, 0.43, 0.6)];
			if (type === 'twoTxTwoObj')
				slots.push(box('body', 3, 0.05, 0.2, 0.43, 0.07), box('body', 4, 0.52, 0.2, 0.43, 0.07));
			break;
		case 'titleOnly':
			slots = [title];
			break;
		case 'objTx':
		case 'picTx':
			slots = [
				title,
				box(type === 'picTx' ? 'pic' : 'obj', 1, 0.36, 0.23, 0.59, 0.64),
				box('body', 2, 0.05, 0.23, 0.27, 0.64),
			];
			break;
		default:
			slots = [title, content];
	}
	return [
		...slots,
		box('dt', 10, 0.05, 0.92, 0.25, 0.04),
		box('ftr', 11, 0.33, 0.92, 0.34, 0.04),
		box('sldNum', 12, 0.75, 0.92, 0.2, 0.04),
	];
}
