import { openZip } from 'ooxml-core/opc';
import { createEditorId, isSelfContainedGlb } from 'ooxml-core/pptx';
import type { PptxElement } from 'ooxml-core/pptx';

import { createInsertDialogFrame } from '../../dialog/insert-dialog-frame';
import type { CanvasSize } from '../types';
import type { InsertText } from './ribbon-graphic-insert';

const GLB_MIME = 'model/gltf-binary';
const MAX_FILE_BYTES = 32 * 1024 * 1024;
const OBJECT_FORMATS = {
	docx: {
		type: 'word',
		part: 'word/document.xml',
		progId: 'Word.Document.12',
		mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	},
	xlsx: {
		type: 'excel',
		part: 'xl/workbook.xml',
		progId: 'Excel.Sheet.12',
		mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	},
	pptx: {
		type: 'powerpoint',
		part: 'ppt/presentation.xml',
		progId: 'PowerPoint.Show.12',
		mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
	},
} as const;

export function pickLocalObject(
	host: HTMLElement,
	canvas: CanvasSize,
	kind: 'model3d' | 'object',
	text: InsertText,
): Promise<PptxElement | undefined> {
	const doc = host.ownerDocument;
	const heading =
		kind === 'model3d'
			? text('pptx.insert.models3d', '3D Models')
			: text('pptx.insert.object', 'Object');
	const body = doc.createElement('div');
	const label = doc.createElement('label');
	const input = doc.createElement('input');
	input.type = 'file';
	input.style.maxWidth = '100%';
	input.accept = kind === 'model3d' ? '.glb' : '.docx,.xlsx,.pptx';
	label.append(text('pptx.backstage.browseDevice', 'Browse this device'), ' ', input);
	const types = doc.createElement('p');
	types.textContent = kind === 'model3d' ? 'GLB' : 'DOCX, XLSX, PPTX';
	const status = doc.createElement('div');
	status.setAttribute('role', 'alert');
	body.append(label, types, status);

	return new Promise((resolve) => {
		let busy = false;
		let closed = false;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => void insert(),
			() => {
				closed = true;
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		function sync() {
			frame.sync(
				heading,
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				text('pptx.common.insert', 'Insert'),
				!busy && !!input.files?.length,
			);
		}
		async function insert() {
			const file = input.files?.[0];
			if (!file || busy) return;
			busy = true;
			status.textContent = '';
			sync();
			try {
				const element = await localElement(doc, file, canvas, kind);
				if (!closed && host.isConnected) {
					resolve(element);
					frame.dialog.close();
				}
			} catch {
				status.textContent = text(
					'pptx.graphic.insertFailed',
					'Could not insert the selected item.',
				);
			} finally {
				busy = false;
				sync();
			}
		}
		input.addEventListener('change', () => {
			status.textContent = '';
			sync();
		});
		sync();
		frame.dialog.show();
	});
}

async function localElement(
	doc: Document,
	file: File,
	canvas: CanvasSize,
	kind: 'model3d' | 'object',
): Promise<PptxElement> {
	if (file.size < 1 || file.size > MAX_FILE_BYTES) throw new Error('Invalid file size');
	const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
	const bytes = await file.arrayBuffer();
	const size = Math.min(canvas.width, canvas.height) * (kind === 'model3d' ? 0.45 : 0.22);
	const box = {
		id: createEditorId(kind),
		name: file.name,
		x: (canvas.width - size) / 2,
		y: (canvas.height - size) / 2,
		width: size,
		height: size,
	};
	const preview = iconPreview(doc, extension.toUpperCase());
	if (kind === 'model3d') {
		if (extension !== 'glb' || !isSelfContainedGlb(new Uint8Array(bytes))) {
			throw new Error('Expected a self-contained GLB 2.0 model');
		}
		return {
			...box,
			type: 'model3d',
			modelMimeType: GLB_MIME,
			modelData: await asDataUrl(bytes, GLB_MIME),
			posterImage: preview,
		};
	}
	if (!Object.hasOwn(OBJECT_FORMATS, extension)) throw new Error('Unsupported object type');
	const format = OBJECT_FORMATS[extension as keyof typeof OBJECT_FORMATS];
	const zip = await openZip(bytes, {
		label: 'Embedded Office object',
		maxInputBytes: MAX_FILE_BYTES,
	});
	if (!zip.file('[Content_Types].xml') || !zip.file(format.part)) {
		throw new Error('The object does not match its file type');
	}
	return {
		...box,
		type: 'ole',
		oleProgId: format.progId,
		oleObjectType: format.type,
		oleName: file.name,
		oleFileExtension: extension,
		oleEmbeddedFileName: file.name,
		oleEmbeddedMimeType: format.mime,
		oleShowAsIcon: true,
		oleEmbeddedData: await asDataUrl(bytes, format.mime),
		previewImageData: preview,
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

function iconPreview(doc: Document, extension: string): string {
	const canvas = doc.createElement('canvas');
	canvas.width = 160;
	canvas.height = 160;
	const ctx = canvas.getContext('2d');
	if (!ctx) throw new Error('Could not render object preview');
	ctx.fillStyle = '#f1f5f9';
	ctx.fillRect(18, 12, 124, 136);
	ctx.strokeStyle = '#334155';
	ctx.lineWidth = 3;
	ctx.strokeRect(18, 12, 124, 136);
	ctx.fillStyle = '#334155';
	ctx.font = 'bold 24px sans-serif';
	ctx.textAlign = 'center';
	ctx.fillText(extension, 80, 88, 110);
	return canvas.toDataURL('image/png');
}
