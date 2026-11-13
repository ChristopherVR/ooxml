import type { PptxElement } from '../../types';

type ImageElement = Extract<PptxElement, { type: 'picture' | 'image' }>;

/** @internal Hydrate requested template pictures in place through the handler's media cache. */
export async function resolveTemplateImages(
	elements: PptxElement[],
	getImageData: (path: string) => Promise<string | undefined>,
): Promise<PptxElement[]> {
	const images = new Map<string, { target: ImageElement; field: 'imageData' | 'svgData' }[]>();
	const pending = [...elements];
	for (const element of pending) {
		if (element.type === 'group') {
			for (const child of element.children) pending.push(child);
		}
		if (element.type !== 'picture' && element.type !== 'image') continue;
		for (const [field, path] of [
			['imageData', element.imagePath],
			['svgData', element.svgPath],
		] as const) {
			if (element[field] || !path || /^(?:https?:|data:|blob:)/u.test(path)) continue;
			const refs = images.get(path) ?? [];
			refs.push({ target: element, field });
			images.set(path, refs);
		}
	}
	for (const [path, refs] of images) {
		const data = await getImageData(path);
		if (data) {
			for (const { target, field } of refs) target[field] = data;
		}
	}
	return elements;
}
