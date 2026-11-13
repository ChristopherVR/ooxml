import type { PptxElement } from '../../types';

function isArchivePath(path: string | undefined): path is string {
	return typeof path === 'string' && path.length > 0 && !/^(?:https?:|data:|blob:)/u.test(path);
}

/** @internal Hydrate requested template assets in place through the handler's media cache. */
export async function resolveTemplateImages(
	elements: PptxElement[],
	getImageData: (path: string) => Promise<string | undefined>,
): Promise<PptxElement[]> {
	const images = new Map<string, { target: object; field: string }[]>();
	const collect = (path: string | undefined, target: object, field: string) => {
		if (!isArchivePath(path)) return;
		const refs = images.get(path) ?? [];
		refs.push({ target, field });
		images.set(path, refs);
	};
	const pending = [...elements];
	for (const element of pending) {
		if (element.type === 'group') {
			for (const child of element.children) pending.push(child);
		}
		if (element.type === 'picture' || element.type === 'image') {
			if (!element.imageData) collect(element.imagePath, element, 'imageData');
			if (!element.svgData) collect(element.svgPath, element, 'svgData');
		}
		if (element.type === 'media' && !element.posterFrameData) {
			collect(element.posterFramePath, element, 'posterFrameData');
		}
		if (element.type === 'model3d') {
			if (!element.modelData) collect(element.modelPath, element, 'modelData');
			if (element.posterImage ? isArchivePath(element.posterImage) : !element.imageData) {
				collect(element.imagePath, element, 'posterImage');
			}
		}
		for (const style of [
			'shapeStyle' in element ? element.shapeStyle : undefined,
			element.type === 'group' ? element.groupFill : undefined,
		]) {
			if (style) collect(style.fillImageUrl, style, 'fillImageUrl');
		}
	}
	for (const [path, refs] of images) {
		const data = await getImageData(path);
		if (data) {
			for (const { target, field } of refs) Object.assign(target, { [field]: data });
		}
	}
	return elements;
}
