import { createEditorId } from 'ooxml-core/pptx';
import type { PptxElement } from 'ooxml-core/pptx';

import { pickLocalModelObject } from '../../form/local-model-object';
import type { CanvasSize } from '../types';
import type { InsertText } from './ribbon-graphic-insert';

export async function pickLocalObject(
	host: HTMLElement,
	canvas: CanvasSize,
	kind: 'model3d' | 'object',
	text: InsertText,
): Promise<PptxElement | undefined> {
	const picked = await pickLocalModelObject(host, kind, {
		heading:
			kind === 'model3d'
				? text('pptx.insert.models3d', '3D Models')
				: text('pptx.insert.object', 'Object'),
		browse: text('pptx.backstage.browseDevice', 'Browse this device'),
		close: text('pptx.common.close', 'Close'),
		cancel: text('pptx.common.cancel', 'Cancel'),
		insert: text('pptx.common.insert', 'Insert'),
		error: text('pptx.graphic.insertFailed', 'Could not insert the selected item.'),
	});
	if (!picked) return undefined;

	const size = Math.min(canvas.width, canvas.height) * (picked.kind === 'model3d' ? 0.45 : 0.22);
	const box = {
		id: createEditorId(picked.kind),
		name: picked.name,
		x: (canvas.width - size) / 2,
		y: (canvas.height - size) / 2,
		width: size,
		height: size,
	};
	if (picked.kind === 'model3d') {
		return {
			...box,
			type: 'model3d',
			modelMimeType: picked.mime,
			modelData: await asDataUrl(picked.bytes, picked.mime),
			posterImage: picked.preview,
		};
	}
	const { format } = picked;
	return {
		...box,
		type: 'ole',
		oleProgId: format.progId,
		oleObjectType: format.type,
		oleName: picked.name,
		oleFileExtension: format.extension,
		oleEmbeddedFileName: picked.name,
		oleEmbeddedMimeType: format.mime,
		oleShowAsIcon: true,
		oleEmbeddedData: await asDataUrl(picked.bytes, format.mime),
		previewImageData: picked.preview,
	};
}

function asDataUrl(bytes: ArrayBuffer, type: string): Promise<string> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = () => resolve(String(reader.result));
		reader.onerror = () => reject(reader.error);
		reader.readAsDataURL(new Blob([bytes], { type }));
	});
}
