import { EditorState, NodeSelection } from 'prosemirror-state';
import { expect, it } from 'vitest';
import { buildArrangeGroup, syncArrange } from './arrange-ribbon';
import { schema } from './schema';

it('enables page and margin alignment for one floating drawing', () => {
	const drawing = schema.node('image', { anchored: true, shape: '{}' });
	const doc = schema.node('doc', null, [schema.node('paragraph', null, [drawing])]);
	const state = EditorState.create({ doc, selection: NodeSelection.create(doc, 1) });
	const group = buildArrangeGroup();
	syncArrange(group, state, true, 'en', true);
	const align = group.querySelector<HTMLSelectElement>('select[data-arrange="align"]');
	expect(align?.disabled).toBe(false);
	expect(
		[...align!.options].filter((option) => option.disabled).map((option) => option.value),
	).toEqual(['selected', 'x', 'y']);
});
