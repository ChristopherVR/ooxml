// Command strings (UI-COMMANDS): ribbon tabs, groups, command labels, menus and dialogs. The
// translations live in the tables of ../../commands/i18n; this file flattens them for 'zh-CN'.
import { commandStrings } from '../../commands/i18n/index';
import { DRAWING_FORMAT_STRINGS } from '../../commands/i18n/drawing-format';
import { flatten } from '../../commands/i18n/types';

export const commands: Record<string, string> = {
	...commandStrings('zh-CN'),
	...flatten([DRAWING_FORMAT_STRINGS], 'zh-CN'),
};
