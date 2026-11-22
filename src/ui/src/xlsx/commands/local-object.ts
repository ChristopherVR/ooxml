import { pictureAnchorAt } from 'ooxml-core/xlsx';
import {
	editing,
	icon,
	sheetLocked,
	target,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { downloadBytes } from '../file-commands';

const INSERTIONS = [
	['model3d', '3D Models', 'insert.3d-models'],
	['object', 'Object', 'insert.object'],
] as const;

async function insertLocal(ctx: EditorContext, kind: 'model3d' | 'object'): Promise<void> {
	const current = target(ctx);
	if (!current) return;
	const { pickLocalModelObject } = await import('../../form/local-model-object');
	const live = () => {
		const active = target(ctx);
		return (
			!ctx.readOnly() &&
			active?.session === current.session &&
			active.ws === current.ws &&
			!sheetLocked(ctx, 'objects')
		);
	};
	if (!live()) return;
	const t = (key: string) => ctx.t(key);
	const error =
		kind === 'model3d'
			? t('Choose a self-contained GLB model.')
			: t('Choose a valid Word, Excel or PowerPoint file.');
	try {
		const item = await pickLocalModelObject(ctx.root, kind, {
			heading: t(kind === 'model3d' ? '3D Models' : 'Insert Object'),
			browse: t('Browse...'),
			close: t('Close'),
			cancel: t('Cancel'),
			insert: t('Insert'),
			error,
		});
		if (!item || !live()) return;
		const bytes = new Uint8Array(item.bytes);
		const png = atob(item.preview.slice('data:image/png;base64,'.length));
		const preview = Uint8Array.from(png, (letter) => letter.charCodeAt(0));
		const size = kind === 'model3d' ? 160 : 96;
		const anchor = pictureAnchorAt(current.active, size, size);
		const drawing = current.session.insertLocalObject(current.sheet, {
			...item,
			bytes,
			preview,
			anchor,
		});
		ctx.selection.set({ drawing });
	} catch {
		ctx.toast(error, 'error');
	}
}

export function selectedEmbeddedObject(ctx: EditorContext) {
	const current = target(ctx);
	const index = ctx.selection.get().drawing;
	const drawing = index === undefined ? undefined : current?.ws.drawings[index];
	return drawing?.kind === 'embeddedObject' ? drawing : undefined;
}

export function localObjectCommands(): Command[] {
	return [
		...INSERTIONS.map(([kind, label, id]) =>
			editing({
				id,
				label,
				icon: icon(kind === 'model3d' ? 'picture' : 'open'),
				lock: 'objects',
				run: (ctx) => insertLocal(ctx, kind),
			}),
		),
		{
			id: 'drawing.download-embedded',
			label: 'Download Embedded File',
			icon: icon('save'),
			enabled: (ctx) => !!selectedEmbeddedObject(ctx),
			run: (ctx) => {
				const drawing = selectedEmbeddedObject(ctx);
				if (!drawing) return;
				const bytes = ctx.session()?.workbook.source?.parts.get(drawing.partName);
				if (!bytes) return;
				const extension = drawing.partName.split('.').pop();
				const name =
					extension && !drawing.name.toLowerCase().endsWith('.' + extension)
						? drawing.name + '.' + extension
						: drawing.name;
				downloadBytes(ctx.host.ownerDocument, bytes, name, drawing.contentType);
			},
		},
	];
}
