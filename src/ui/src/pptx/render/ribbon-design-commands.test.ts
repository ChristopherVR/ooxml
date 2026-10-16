import { describe, expect, it } from 'vitest';

import {
	DESIGN_RIBBON_COMMANDS,
	DESIGN_RIBBON_GROUPS,
	designCommandState,
} from './ribbon-design-commands';

describe('design command contract', () => {
	it('places five unique commands in the themes, customize and designer groups', () => {
		expect(new Set(DESIGN_RIBBON_COMMANDS.map((command) => command.id)).size).toBe(5);
		expect(
			DESIGN_RIBBON_COMMANDS.every((command) =>
				DESIGN_RIBBON_GROUPS.some((group) => command.id.startsWith(`${group.id}.`)),
			),
		).toBeTruthy();
	});

	it('locks edits but preserves slide-size inspection in read-only mode', () => {
		for (const command of DESIGN_RIBBON_COMMANDS) {
			expect(designCommandState(command.id, { editable: false }).disabled).toBe(
				command.id !== 'design.customize.slideSize',
			);
		}
	});

	it('lists Designer but never enables it', () => {
		for (const editable of [true, false]) {
			expect(designCommandState('design.designer.designIdeas', { editable }).disabled).toBe(true);
		}
	});

	it('reflects host open state and optional action availability', () => {
		expect(
			designCommandState('design.themes.editTheme', { editable: true, editorOpen: true }),
		).toMatchObject({ active: true, expanded: true });
		expect(
			designCommandState('design.customize.slideSize', { editable: true, hasSlideSize: false })
				.hidden,
		).toBeTruthy();
		expect(
			designCommandState('design.customize.formatBackground', {
				editable: true,
				hasBackground: false,
			}).hidden,
		).toBeTruthy();
	});
});
