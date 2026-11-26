// Command strings (UI-COMMANDS): ribbon tabs, groups, command labels, menus and dialogs. The
// translations live in the tables of ../../commands/i18n; this file flattens them for 'en'.
import { commandStrings } from '../../commands/i18n/index';
import { DRAWING_FORMAT_KEYS } from '../../commands/i18n/drawing-format-keys';

export const commands: Record<string, string> = {
	...commandStrings('en'),
	...Object.fromEntries(DRAWING_FORMAT_KEYS.map((key) => [key, key])),
};
