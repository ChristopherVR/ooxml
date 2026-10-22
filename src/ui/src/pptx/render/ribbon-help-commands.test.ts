import { describe, expect, it } from 'vitest';

import { buildHelpRibbon, HELP_RIBBON_COMMANDS } from './ribbon-help-commands';

const t = (key: string) => `t:${key}`;
const ids = (state: Parameters<typeof buildHelpRibbon>[1]) =>
	buildHelpRibbon(t, state)[0]!.commands.map((command) => command.id);

describe('help ribbon model', () => {
	it('builds one group of compact, translated commands in the shared order', () => {
		const [group, ...rest] = buildHelpRibbon(t, { options: true });
		expect(rest).toHaveLength(0);
		expect(group).toMatchObject({ id: 'help.help', label: 't:pptx.ribbon.tab.help' });
		expect(group!.commands).toStrictEqual(
			HELP_RIBBON_COMMANDS.map((command) => ({
				id: command.id,
				label: `t:${command.labelKey}`,
				icon: command.icon,
				compact: true,
			})),
		);
	});

	it('leaves out Options when the dialog is unavailable, and Shortcuts only when asked', () => {
		expect(ids({ options: false })).toStrictEqual([
			'help.help.keyboardShortcuts',
			'help.help.accessibility',
		]);
		expect(ids({ options: false, shortcuts: false })).toStrictEqual(['help.help.accessibility']);
		expect(ids({ options: true, shortcuts: true })).toStrictEqual([
			'help.help.options',
			'help.help.keyboardShortcuts',
			'help.help.accessibility',
		]);
	});
});
