import type { PendingMediaPart } from 'ooxml-core/docx';
import { GLB_PART_CONTENT_TYPE } from 'ooxml-core/drawingml';
import type { EditorView } from 'prosemirror-view';
import { focusView } from './focus-view';
import type { InsertControllerHost } from './insert-controller';
import { translate, type EditorLocale } from './localization';
import { insertPicture, newPicturePartName } from './picture-commands';

/** Insert at the selection that requested the picker, after its files have been validated. */
export async function insertLocalDrawing(
	view: EditorView,
	kind: 'model3d' | 'object',
	locale: EditorLocale,
	host: InsertControllerHost,
	pendingMedia: Map<string, PendingMediaPart>,
): Promise<void> {
	const parent = host.toolbar();
	if (!parent) return;
	const { doc, selection } = view.state;
	const current = () =>
		host.view() === view &&
		view.editable &&
		view.state.doc === doc &&
		view.state.selection.eq(selection);
	const t = (key: Parameters<typeof translate>[1]) => translate(locale, key);
	try {
		const { pickLocalModelObject } = await import('../form/local-model-object');
		if (!parent.isConnected || !current()) return;
		const item = await pickLocalModelObject(parent, kind, {
			heading: t(kind === 'model3d' ? '3D Models' : 'Object'),
			browse: t('Browse…'),
			close: t('Close'),
			cancel: t('Cancel'),
			insert: t('Insert'),
			error: t('Could not insert the selected file.'),
		});
		if (!item || !current()) return;
		const previewType = 'image/png';
		const previewName = (host.newMediaPartName ?? newPicturePartName)(previewType);
		const stem = previewName.slice(0, -4);
		const payload =
			item.kind === 'model3d'
				? { kind: 'model3d' as const, partName: `${stem}.glb` }
				: {
						kind: 'officePackage' as const,
						partName: `${stem.replace(/^word\/media\//, 'word/embeddings/')}.${item.format.extension}`,
						progId: item.format.progId,
					};
		const previewBase64 = item.preview.split(',')[1];
		if (!previewBase64) throw new Error('Missing preview bytes');
		const previewBytes = Uint8Array.from(atob(previewBase64), (character) =>
			character.charCodeAt(0),
		);
		const stage = (name: string, part: PendingMediaPart) => {
			host.stageMedia?.(name, part);
			pendingMedia.set(name, part);
		};
		stage(payload.partName, {
			bytes: new Uint8Array(item.bytes),
			contentType: item.kind === 'model3d' ? GLB_PART_CONTENT_TYPE : item.format.mime,
		});
		stage(previewName, { bytes: previewBytes, contentType: previewType });
		const size = Math.min(Math.max(16, host.contentWidth()), item.kind === 'model3d' ? 192 : 96);
		insertPicture(view, {
			partName: previewName,
			relId: '',
			contentType: previewType,
			payload,
			widthPx: size,
			heightPx: size,
			altText: item.name,
			title: item.name,
		});
		focusView(view);
	} catch {
		host.reportError(new Error(t('Could not insert the selected file.')));
	}
}
