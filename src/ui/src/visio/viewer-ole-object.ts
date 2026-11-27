import {
	isOpenableVisioOleObject,
	readVisioEmbeddedObject,
	type VisioOleObject,
	type VisioShape,
} from 'ooxml-core/visio';
import { FILE_COMMAND_EVENT } from 'ooxml-core/embedding';
import type { ViewerController } from './controller';

/** The detail of the cancelable `file-command` event Open on an embedded document raises. */
export interface OpenEmbeddedDocumentDetail {
	command: 'openEmbeddedDocument';
	fileName: string;
	bytes: Uint8Array;
}

const MIME: Record<string, string> = {
	docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
	vsdx: 'application/vnd.ms-visio.drawing',
};

/** The shape `id` when it holds a general embedded OLE object, looking inside groups. */
export function oleObjectShape(shapes: readonly VisioShape[], id: string): VisioShape | undefined {
	for (const shape of shapes) {
		if (shape.id === id) return shape.oleObject ? shape : undefined;
		const inner = oleObjectShape(shape.children, id);
		if (inner) return inner;
	}
	return undefined;
}

/**
 * Open on an embedded object (double-click, or the shape menu): an embedded Word, Excel,
 * PowerPoint or Visio file is announced as the cancelable `file-command` event
 * `openEmbeddedDocument`, so a host such as the Office suite opens it in its own editor; when no
 * host takes it, it is downloaded. Any other object is refused with the reason. Opening reads the
 * current drawing and never changes it; the object stays embedded as it is.
 */
export async function openOleObject(
	host: HTMLElement,
	controller: Pick<ViewerController, 'exportVsdx'>,
	object: VisioOleObject,
	announce: (message: string) => void,
): Promise<void> {
	if (!isOpenableVisioOleObject(object)) {
		announce(object.refusal ?? 'This embedded object cannot be opened here.');
		return;
	}
	try {
		const { bytes, fileName } = await readVisioEmbeddedObject(
			controller.exportVsdx().bytes,
			object.part,
		);
		const detail: OpenEmbeddedDocumentDetail = {
			command: 'openEmbeddedDocument',
			fileName,
			bytes,
		};
		const unhandled = host.dispatchEvent(
			new CustomEvent(FILE_COMMAND_EVENT, {
				detail,
				bubbles: true,
				composed: true,
				cancelable: true,
			}),
		);
		if (unhandled) download(host, bytes, fileName, MIME[object.format]);
	} catch (error) {
		announce(error instanceof Error ? error.message : 'The embedded object could not be opened.');
	}
}

function download(host: HTMLElement, bytes: Uint8Array, name: string, type?: string): void {
	const doc = host.ownerDocument;
	const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], type ? { type } : undefined));
	const anchor = doc.createElement('a');
	anchor.href = url;
	anchor.download = name;
	anchor.hidden = true;
	(host.shadowRoot ?? host).append(anchor);
	anchor.click();
	anchor.remove();
	doc.defaultView?.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
