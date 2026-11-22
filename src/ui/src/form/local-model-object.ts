import { GLB_MIME, isSelfContainedGlb } from 'ooxml-core/drawingml';
import {
	MAX_LOCAL_OBJECT_BYTES,
	readEmbeddableOfficePackage,
	type EmbeddableOfficePackage,
} from 'ooxml-core/opc';

import { withTokens } from '../base';
import { createInsertDialogFrame } from '../dialog/insert-dialog-frame';
import css from './local-model-object.css?raw';

const filePickerStyles = withTokens(css);

type LocalObjectKind = 'model3d' | 'object';

interface LocalObjectFile {
	name: string;
	bytes: ArrayBuffer;
	/** A PNG icon for use in the product and as an unsupported-renderer fallback. */
	preview: string;
}

export type LocalModelObject = LocalObjectFile &
	(
		| { kind: 'model3d'; mime: typeof GLB_MIME }
		| { kind: 'object'; format: EmbeddableOfficePackage }
	);

interface LocalObjectLabels {
	heading: string;
	browse: string;
	close: string;
	cancel: string;
	insert: string;
	error: string;
}

/** Pick local files without fetching their contents or accepting external model resources. */
export function pickLocalModelObject(
	host: HTMLElement | ShadowRoot,
	kind: LocalObjectKind,
	labels: LocalObjectLabels,
): Promise<LocalModelObject | undefined> {
	const doc = host.ownerDocument;
	const body = doc.createElement('div');
	body.className = 'office-local-object-picker';
	const style = doc.createElement('style');
	style.textContent = filePickerStyles;

	const browse = doc.createElement('button');
	browse.type = 'button';
	browse.textContent = labels.browse;

	const input = doc.createElement('input');
	input.type = 'file';
	input.hidden = true;
	input.setAttribute('aria-label', labels.browse);
	input.accept = kind === 'model3d' ? '.glb' : '.docx,.xlsx,.pptx';
	browse.addEventListener('click', () => input.click());

	const types = doc.createElement('p');
	types.textContent = kind === 'model3d' ? 'GLB' : 'DOCX, XLSX, PPTX';
	const filename = doc.createElement('div');
	filename.setAttribute('role', 'status');
	const status = doc.createElement('div');
	status.setAttribute('role', 'alert');
	body.append(style, browse, input, types, filename, status);

	return new Promise((resolve) => {
		let busy = false;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => void insert(),
			() => {
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		function sync() {
			browse.disabled = busy;
			frame.sync(
				labels.heading,
				labels.close,
				labels.cancel,
				labels.insert,
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
				const item = await localObject(doc, file, kind);
				if (frame.dialog.open && frame.dialog.isConnected) {
					resolve(item);
					frame.dialog.close();
				}
			} catch {
				status.textContent = labels.error;
			} finally {
				busy = false;
				sync();
				if (!frame.dialog.isConnected) resolve(undefined);
			}
		}
		input.addEventListener('change', () => {
			filename.textContent = input.files?.[0]?.name ?? '';
			status.textContent = '';
			sync();
		});
		sync();
		frame.dialog.show();
	});
}

async function localObject(
	doc: Document,
	file: File,
	kind: LocalObjectKind,
): Promise<LocalModelObject> {
	if (file.size < 1 || file.size > MAX_LOCAL_OBJECT_BYTES) throw new Error('Invalid file size');
	const extension = file.name.split('.').pop()?.toLowerCase() ?? '';
	const bytes = await file.arrayBuffer();
	if (kind === 'model3d') {
		if (extension !== 'glb' || !isSelfContainedGlb(new Uint8Array(bytes))) {
			throw new Error('Expected a self-contained GLB 2.0 model');
		}
		return { kind, name: file.name, bytes, mime: GLB_MIME, preview: iconPreview(doc, 'GLB') };
	}
	const format = await readEmbeddableOfficePackage(bytes, extension);
	return {
		kind,
		name: file.name,
		bytes,
		format,
		preview: iconPreview(doc, format.extension.toUpperCase()),
	};
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
