import { VisioPackageError } from './package.js';

/** Metadata has separate limits from document text and geometry formula content. */
export function metadata(value: string, max: number, label: string): string {
	if (value.length > max)
		throw new VisioPackageError('METADATA_LIMIT', `${label} exceeds its metadata length limit.`);
	return value;
}
export function metadataAttributes(node: Element): Map<string, string> {
	return new Map(
		Array.from(node.attributes).map((item) => {
			const max = ['Name', 'NameU'].includes(item.name) ? 4096 : 256;
			return [
				metadata(item.name, 256, 'Attribute name'),
				metadata(item.value, max, 'Shape/style attribute'),
			];
		}),
	);
}
