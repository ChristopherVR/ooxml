import { IMAGE_EXTENSIONS } from 'ooxml-core/xlsx';
import { editing, icon, target, type Command } from 'ooxml-core/xlsx/ui';
import { pickImage } from './insert';

export function sheetPictureCommands(): Command[] {
	return [
		editing({
			id: 'page.background',
			label: 'Background',
			icon: icon('background'),
			lock: 'objects',
			run: async (ctx) => {
				const t = target(ctx);
				if (!t) return;
				const file = await pickImage(ctx);
				if (!file) return;
				if (!Object.hasOwn(IMAGE_EXTENSIONS, file.type) || !file.size) {
					ctx.toast(
						ctx.t('This picture format is not supported. Use PNG, JPEG, GIF, BMP or WebP.'),
						'warning',
					);
					return;
				}
				const bytes = new Uint8Array(await file.arrayBuffer());
				const live = target(ctx);
				if (
					live?.session !== t.session ||
					live.ws !== t.ws ||
					!ctx.commands.isEnabled('page.background')
				)
					return;
				t.session.setSheetBackground(live.sheet, { bytes, contentType: file.type });
			},
		}),
		editing({
			id: 'page.background-delete',
			label: 'Delete Background',
			icon: icon('background'),
			lock: 'objects',
			enabled: (ctx) => {
				const sheet = target(ctx)?.ws;
				return !!(sheet?.backgroundPicture || sheet?.preserved.has('picture'));
			},
			run: (ctx) => {
				const t = target(ctx);
				if (t) t.session.setSheetBackground(t.sheet, undefined);
			},
		}),
	];
}
