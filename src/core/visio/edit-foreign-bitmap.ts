import { attribute, children } from './sheet';

/**
 * A local embedded bitmap (Type="Foreign" with one Bitmap ForeignData relationship and nothing
 * else foreign). Its placement is entirely ShapeSheet cells, so the same move, resize and delete
 * admission as an ordinary 2D shape applies; the media relationship is left untouched.
 */
export function isLocalBitmapShape(shape: Element): boolean {
	const data = children(shape, 'ForeignData');
	return (
		attribute(shape, 'Type') === 'Foreign' &&
		!shape.hasAttribute('Master') &&
		!shape.hasAttribute('MasterShape') &&
		!children(shape, 'Shapes').length &&
		!children(shape, 'Rel').length &&
		data.length === 1 &&
		attribute(data[0], 'ForeignType') === 'Bitmap' &&
		children(data[0], 'Rel').length === 1 &&
		!Array.from(data[0]!.childNodes).some(
			(node) =>
				(node.nodeType === 1 && (node as Element).localName !== 'Rel') ||
				(node.nodeType === 3 && node.textContent!.trim() !== ''),
		)
	);
}
