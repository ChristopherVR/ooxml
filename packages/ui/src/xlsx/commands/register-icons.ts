import { registerIcon } from '../../index';
import { SHEET_ICONS } from 'ooxml-core/xlsx/ui';

let registered = false;

/** Registers every glyph once (idempotent; safe during SSR because it touches no DOM). */
export function registerSheetIcons(): void {
	if (registered) return;
	registered = true;
	for (const [name, d] of Object.entries(SHEET_ICONS)) registerIcon(`xl-${name}`, d);
}
