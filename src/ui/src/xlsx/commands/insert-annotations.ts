import { editing, icon, target, type Command } from 'ooxml-core/xlsx/ui';
import { requestHeaderEdit } from '../page-views/header-edit-request';
import { setViewMode } from './view';

export function insertAnnotationCommands(): Command[] {
	return [
		editing({
			id: 'insert.note',
			label: 'New Note',
			icon: icon('commentNew'),
			lock: 'objects',
			enabled: (ctx) => {
				const t = target(ctx);
				return (
					!!t &&
					!t.ws.comments.some(
						({ address, replies }) =>
							replies && address.row === t.active.row && address.col === t.active.col,
					)
				);
			},
			run: async (ctx) => void (await ctx.dialogs.open('note')),
		}),
		editing({
			id: 'insert.header-footer',
			label: 'Header & Footer',
			icon: icon('header'),
			lock: false,
			run: (ctx) => {
				const sheet = target(ctx)?.ws;
				if (!sheet) return;
				requestHeaderEdit(ctx, sheet);
				setViewMode(ctx, 'pageLayout');
			},
		}),
	];
}
