// Insert tab: Table, Pictures, Charts, Link, Comment and Symbol.
import { IMAGE_EXTENSIONS, pictureAnchorAt, type ChartType } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, target } from './util.js';

/** Picture types the core stores that a browser can also paint. */
const PICTURE_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp'].filter(
	(type) => type in IMAGE_EXTENSIONS,
);

export const CHART_TYPES: ReadonlyArray<readonly [type: ChartType, label: string, glyph: string]> =
	[
		['column', 'Column', 'chartColumn'],
		['line', 'Line', 'chartLine'],
		['pie', 'Pie', 'chartPie'],
		['bar', 'Bar', 'chartBar'],
		['area', 'Area', 'chartArea'],
		['scatter', 'Scatter', 'chartScatter'],
		['doughnut', 'Doughnut', 'chartPie'],
		['radar', 'Radar', 'chartLine'],
	];

/** Picks an image file with a native file input. Resolves undefined when cancelled. */
function pickImage(ctx: EditorContext): Promise<File | undefined> {
	return new Promise((resolve) => {
		const input = ctx.host.ownerDocument.createElement('input');
		input.type = 'file';
		input.accept = PICTURE_TYPES.join(',');
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
export async function insertPicture(ctx: EditorContext, file: File): Promise<void> {
	const t = target(ctx);
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
	t.session.addImage(
		t.sheet,
		bytes,
		file.type,
		pictureAnchorAt(t.active, width, height),
		file.name,
	);
	ctx.selection.set({ drawing: t.ws.drawings.length - 1 });
}

export function insertCommands(): Command[] {
	return [
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
			label: 'Pictures',
			icon: icon('picture'),
			lock: 'objects',
			run: async (ctx) => {
				const file = await pickImage(ctx);
				if (file) await insertPicture(ctx, file);
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
