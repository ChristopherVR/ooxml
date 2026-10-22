import { describe, expect, it } from 'vitest';

import { buildRecordRibbon, RECORD_COMMAND_GROUPS } from './ribbon-record-commands';
import { ribbonGroupViews, SLIDE_SHOW_COMMAND_GROUPS } from './ribbon-slide-show-commands';

const t = (key: string) => `t:${key}`;

describe('record ribbon model', () => {
	it('translates the four groups and disables the unsupported placeholders', () => {
		const groups = buildRecordRibbon(t);
		expect(groups.map((group) => group.id)).toStrictEqual(
			RECORD_COMMAND_GROUPS.map((group) => group.id),
		);
		const commands = groups.flatMap((group) => group.commands);
		expect(commands).toHaveLength(6);
		expect(
			commands.filter((command) => command.disabled).map((command) => command.id),
		).toStrictEqual([
			'record.camera.cameo',
			'record.manage.clear',
			'record.manage.reset',
			'record.help.learnMore',
		]);
		expect(commands.find((command) => command.id === 'record.record.fromBeginning')).toStrictEqual({
			id: 'record.record.fromBeginning',
			label: 't:pptx.slideShow.fromBeginning',
			icon: 'video',
		});
		for (const group of groups) {
			for (const command of group.commands) {
				expect(command.id.startsWith(`${group.id}.`)).toBeTruthy();
			}
		}
	});

	it('turns a catalogue tooltip key into the command ScreenTip', () => {
		const [start] = ribbonGroupViews(t, SLIDE_SHOW_COMMAND_GROUPS);
		expect(start!.commands[0]).toMatchObject({
			id: 'slideShow.startSlideShow.fromBeginning',
			title: 't:pptx.slideShow.fromBeginningTooltip',
		});
	});
});
