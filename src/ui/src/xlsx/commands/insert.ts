// Insert tab: Table, Pictures, Charts, Link, Comment and Symbol.
import { canCaptureDisplay } from '../../form/display-capture';
import { captureScreenshotFile } from '../../form/display-clipping';
import { getCell, IMAGE_EXTENSIONS, pictureAnchorAt, type ChartType } from 'ooxml-core/xlsx';
import { pickMicrosoftForm } from '../../form/microsoft-forms';
import { openOfficeAddIns } from '../../ribbon/add-in-dialog';
import { pickOfficeIcon } from '../../form/icon-picker';
import { iconPngFile } from '../../form/local-svg';
import type { Command } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { icon } from 'ooxml-core/xlsx/ui';
import { editing, target, sheetLocked, INSERT_CHART_TYPES } from 'ooxml-core/xlsx/ui';
import {
	cellPictureCommands,
	cellPictureEditsAvailable,
	cellPictureUnavailable,
	insertCellPicture,
} from './cell-pictures';

/** Picture types the core stores that a browser can also paint. */
const PICTURE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp'].filter(
	(type) => type in IMAGE_EXTENSIONS,
);

const CHART_GLYPHS: Partial<Record<ChartType, string>> = {
	column: 'chartColumn',
	line: 'chartLine',
	pie: 'chartPie',
	bar: 'chartBar',
	area: 'chartArea',
	scatter: 'chartScatter',
	doughnut: 'chartPie',
	radar: 'chartLine',
};

/** The chart types of Insert Chart (shared with other products) with the ribbon glyph of each. */
export const CHART_TYPES: ReadonlyArray<readonly [type: ChartType, label: string, glyph: string]> =
	INSERT_CHART_TYPES.map(([type, label]) => [type, label, CHART_GLYPHS[type] ?? 'chartColumn']);

/** Picks an image file with a native file input. Resolves undefined when cancelled. */
export function pickImage(
	ctx: EditorContext,
	accept = PICTURE_TYPES.join(','),
): Promise<File | undefined> {
	return new Promise((resolve) => {
		const input = ctx.host.ownerDocument.createElement('input');
		input.type = 'file';
		input.accept = accept;
		input.hidden = true;
		input.addEventListener('change', () => {
			resolve(input.files?.[0] ?? undefined);
			input.remove();
		});
		input.addEventListener('cancel', () => {
			resolve(undefined);
			input.remove();
		});
		ctx.root.append(input);
		input.click();
	});
}

async function imageSize(file: Blob): Promise<{ width: number; height: number }> {
	try {
		const bitmap = await createImageBitmap(file);
		const size = { width: bitmap.width, height: bitmap.height };
		bitmap.close();
		return size;
	} catch {
		return { width: 320, height: 240 };
	}
}

/** Inserts a picture file at the active cell (scaled to at most 640 px wide, like Excel's fit). */
export async function insertPicture(
	ctx: EditorContext,
	file: File,
	t = target(ctx),
): Promise<void> {
	if (!t) return;
	if (!PICTURE_TYPES.includes(file.type)) {
		ctx.toast(
			ctx.t('This picture format is not supported. Use PNG, JPEG, GIF, BMP or WebP.'),
			'warning',
		);
		return;
	}
	const bytes = new Uint8Array(await file.arrayBuffer());
	const { width, height } = await imageSize(file);
	const live = target(ctx);
	if (ctx.readOnly() || live?.session !== t.session || live.ws !== t.ws) return;
	t.session.addImage(
		t.sheet,
		bytes,
		file.type,
		pictureAnchorAt(t.active, width, height),
		file.name,
	);
	ctx.selection.set({ drawing: t.ws.drawings.length - 1 });
}

async function wordArtDialog(ctx: EditorContext, edit: boolean): Promise<void> {
	const current = target(ctx);
	if (!current) return;
	const index = ctx.selection.get().drawing;
	const drawing = edit && index !== undefined ? current.ws.drawings[index] : undefined;
	const shape = drawing?.kind === 'unsupported' ? drawing.shape : undefined;
	const initial =
		shape?.wordArt && shape.text ? { text: shape.text.text, style: shape.wordArt } : undefined;
	if (edit && !initial) return;
	const t = (label: string) => ctx.t(label);
	const { pickWordArt } = await import('../../form/word-art-picker');
	if (ctx.session() !== current.session || ctx.activeSheet() !== current.sheet || ctx.readOnly())
		return;
	const value = await pickWordArt(
		ctx.root,
		{
			heading: edit ? t('Edit WordArt') : t('WordArt'),
			text: t('WordArt text'),
			placeholder: t('Your text here'),
			style: (n) => t('WordArt style {n}').replace('{n}', String(n)),
			transform: t('Transform'),
			preview: t('Preview'),
			locale: ctx.locale(),
			confirm: edit ? t('Apply') : t('Insert'),
			cancel: t('Cancel'),
			close: t('Close'),
		},
		initial,
	);
	const live = target(ctx);
	if (!value || ctx.readOnly() || current.session !== live?.session || current.ws !== live.ws)
		return;
	if (edit) {
		if (index !== undefined && drawing === live.ws.drawings[index])
			live.session.editWordArt(live.sheet, index, value);
		return;
	}
	const anchor = pictureAnchorAt(
		current.active,
		Math.min(650, Math.max(190, value.text.length * 23 + 24)),
		value.style.warp ? 115 : 75,
	);
	const added = live.session.addWordArt(live.sheet, anchor, value);
	ctx.selection.set({ drawing: added });
}

export function selectedSvg(ctx: EditorContext) {
	const current = target(ctx);
	const index = ctx.selection.get().drawing;
	if (!current || index === undefined) return undefined;
	const drawing = current.ws.drawings[index];
	return drawing?.kind === 'image' && drawing.svgPartName ? { current, drawing, index } : undefined;
}

async function graphicsFill(ctx: EditorContext): Promise<void> {
	const selected = selectedSvg(ctx);
	if (!selected) return;
	const { pickGraphicsFill } = await import('../../form/graphics-fill');
	const t = (label: string) => ctx.t(label);
	const patch = await pickGraphicsFill(ctx.root, selected.drawing.effects, {
		heading: t('Graphics Fill'),
		color: t('Color'),
		reset: t('No Recolor'),
		cancel: t('Cancel'),
		close: t('Close'),
		apply: t('Apply'),
	});
	const live = selectedSvg(ctx);
	if (
		!patch ||
		ctx.readOnly() ||
		live?.current.session !== selected.current.session ||
		live.drawing !== selected.drawing ||
		live.index !== selected.index
	)
		return;
	live.current.session.editPictureEffects(live.current.sheet, live.index, patch);
}

export function insertCommands(): Command[] {
	return [
		{
			id: 'insert.get-add-ins',
			label: 'Get Add-ins',
			icon: 'addIns',
			run: (ctx) => {
				void openOfficeAddIns(
					ctx.root,
					'getAddIns',
					(label) => ctx.t(label),
					'excel',
					ctx.locale(),
				);
			},
		},
		{
			id: 'insert.my-add-ins',
			label: 'My Add-ins',
			icon: 'addIns',
			run: (ctx) => {
				void openOfficeAddIns(ctx.root, 'myAddIns', (label) => ctx.t(label), 'excel', ctx.locale());
			},
		},
		editing({
			id: 'insert.forms',
			label: 'Forms',
			icon: 'forms',
			lock: 'insertHyperlinks',
			run: async (ctx) => {
				const current = target(ctx);
				if (!current) return;
				const form = await pickMicrosoftForm(ctx.root, (label) => ctx.t(label));
				const live = target(ctx);
				if (
					!form ||
					ctx.readOnly() ||
					sheetLocked(ctx, 'insertHyperlinks') ||
					live?.ws !== current.ws ||
					live.session !== current.session
				)
					return;
				const { row, col } = current.active;
				const range = { start: current.active, end: current.active };
				const cell = getCell(current.ws, row, col);
				current.session.batch('Insert form link', () => {
					if (!cell?.formula && (cell?.value == null || cell.value === ''))
						current.session.setCellValue(current.sheet, row, col, form.title);
					current.session.setHyperlink(current.sheet, range, {
						target: form.url,
						display: form.title,
					});
				});
			},
		}),
		editing({
			id: 'insert.table',
			label: 'Table',
			icon: icon('table'),
			shortcut: 'Ctrl+T',
			lock: false,
			run: (ctx) => void ctx.dialogs.open('create-table'),
		}),
		editing({
			id: 'insert.pictures',
			label: 'Place over Cells',
			icon: icon('picture'),
			lock: 'objects',
			run: async (ctx) => {
				const file = await pickImage(ctx);
				if (file) await insertPicture(ctx, file);
			},
		}),
		editing({
			id: 'insert.picture-in-cell',
			label: 'Place in Cell',
			unavailable: cellPictureUnavailable('Select a cell on an editable sheet.'),
			enabled: cellPictureEditsAvailable,
			icon: icon('picture'),
			run: async (ctx) => {
				const current = target(ctx);
				if (!current) return;
				const file = await pickImage(ctx);
				if (file) await insertCellPicture(ctx, current, file);
			},
		}),
		...cellPictureCommands(),
		...(
			[
				['screenshot', 'Screenshot'],
				['screen-clipping', 'Screen Clipping'],
			] as const
		).map(([id, label]) =>
			editing({
				id: 'insert.' + id,
				label,
				icon: icon('picture'),
				lock: 'objects',
				enabled: canCaptureDisplay,
				run: async (ctx) => {
					const current = target(ctx);
					if (!current) return;
					const t = (key: string) => ctx.t(key);
					const file = await captureScreenshotFile(
						ctx.root,
						t('Screenshot'),
						id === 'screen-clipping'
							? {
									heading: t('Screen Clipping'),
									close: t('Close'),
									cancel: t('Cancel'),
									insert: t('Insert'),
									x: 'X',
									y: 'Y',
									width: t('Width'),
									height: t('Height'),
								}
							: undefined,
					);
					if (file) await insertPicture(ctx, file, current);
				},
			}),
		),
		editing({
			id: 'insert.icons',
			label: 'Icons',
			icon: 'insertIcons',
			lock: 'objects',
			run: async (ctx) => {
				const before = target(ctx);
				if (!before) return;
				const t = (key: string) => ctx.t(key);
				try {
					const selected = await pickOfficeIcon(ctx.root, {
						heading: t('Icons'),
						search: t('Search icons'),
						color: t('Icon color'),
						empty: t('No matching icons.'),
						cancel: t('Cancel'),
						insert: t('Insert'),
						close: t('Close'),
						iconName: t,
					});
					if (!selected.length) return;
					const icons = await Promise.all(
						selected.map(async (icon) => {
							const file = await iconPngFile(ctx.host.ownerDocument, icon);
							return {
								name: icon.name,
								png: new Uint8Array(await file.arrayBuffer()),
								svg: new TextEncoder().encode(icon.svg),
							};
						}),
					);
					const live = target(ctx);
					if (live?.session !== before.session || live.ws !== before.ws || ctx.readOnly()) return;
					live.session.batch(t('Icons'), () => {
						for (const [index, icon] of icons.entries()) {
							const anchor = pictureAnchorAt(before.active, 96, 96);
							anchor.from.colOffset += (index % 4) * 104 * 9525;
							anchor.from.rowOffset += Math.floor(index / 4) * 104 * 9525;
							live.session.addImage(live.sheet, icon.png, 'image/png', anchor, icon.name, icon.svg);
						}
					});
					ctx.selection.set({ drawing: live.ws.drawings.length - 1 });
				} catch {
					ctx.toast(t('Could not insert the icon.'), 'error');
				}
			},
		}),
		editing({
			id: 'insert.chart',
			label: 'Recommended Charts',
			icon: icon('chartRecommended'),
			lock: 'objects',
			run: (ctx, arg) =>
				void ctx.dialogs.open('insert-chart', { type: typeof arg === 'string' ? arg : 'column' }),
		}),
		editing({
			id: 'insert.word-art',
			label: 'WordArt',
			icon: 'insertWordArt',
			lock: 'objects',
			run: (ctx) => wordArtDialog(ctx, false),
		}),
		editing({
			id: 'drawing.graphics-fill',
			label: 'Graphics Fill',
			icon: icon('fillColor'),
			lock: 'objects',
			run: (ctx) => graphicsFill(ctx).catch((error: unknown) => ctx.toast(String(error), 'error')),
		}),
		editing({
			id: 'drawing.edit-word-art',
			label: 'Edit WordArt',
			icon: 'insertWordArt',
			lock: 'objects',
			run: (ctx) => wordArtDialog(ctx, true),
		}),
		...CHART_TYPES.slice(0, 6).map(([type, label, glyph]) =>
			editing({
				id: `insert.chart-${type}`,
				label: `Insert ${label} Chart`,
				icon: icon(glyph),
				lock: 'objects',
				run: (ctx) => void ctx.dialogs.open('insert-chart', { type }),
			}),
		),
		editing({
			id: 'insert.link',
			label: 'Link',
			icon: icon('link'),
			shortcut: 'Ctrl+K',
			lock: false,
			run: (ctx) => void ctx.dialogs.open('hyperlink'),
		}),
		editing({
			id: 'insert.comment',
			label: 'Comment',
			icon: icon('commentNew'),
			lock: false,
			run: (ctx) => void ctx.commands.run('review.new-comment'),
		}),
		editing({
			id: 'insert.symbol',
			label: 'Symbol',
			icon: icon('symbol'),
			lock: false,
			run: (ctx) => void ctx.dialogs.open('symbol'),
		}),
	];
}
