// Format Cells > Protection: Locked and Hidden (effective once the sheet is protected).
import { checkbox, el, text } from '../fields.js';
import type { FormatTab, TabInit } from './types.js';

export function protectionTab(init: TabInit): FormatTab {
	const { ctx } = init;
	const p = init.style.protection ?? {};
	const locked = checkbox(ctx, 'Locked', p.locked !== false);
	const hidden = checkbox(ctx, 'Hidden', !!p.hidden);
	const changed = new Set<'locked' | 'hidden'>();
	locked.input.addEventListener('change', () => changed.add('locked'));
	hidden.input.addEventListener('change', () => changed.add('hidden'));
	const panel = el(ctx, 'div');
	panel.append(
		locked.wrapper,
		hidden.wrapper,
		text(
			ctx,
			'Locking cells or hiding formulas has no effect until you protect the worksheet (Review tab, Protect Sheet).',
		),
	);
	return {
		id: 'protection',
		label: 'Protection',
		panel,
		dirty: () => changed.size > 0,
		patch: () => {
			const protection: { locked?: boolean; hidden?: boolean } = {};
			if (changed.has('locked')) protection.locked = locked.input.checked;
			if (changed.has('hidden')) protection.hidden = hidden.input.checked;
			return changed.size ? { protection } : {};
		},
		focus: () => locked.input.focus(),
	};
}
