// Command strings (UI-COMMANDS): ribbon tabs, groups, command labels, menus and dialogs. The
// translations live in the tables of ../../commands/i18n; this file flattens them for 'zh-CN'.
import { commandStrings } from '../../commands/i18n/index.js';

export const commands: Record<string, string> = commandStrings('zh-CN');
