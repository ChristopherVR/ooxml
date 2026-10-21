// Which of a run's tabs Word stores as positional tabs (`w:ptab`) is part of the run's content,
// though the editor carries it with the run's other properties. Text that comes in as text, pasted
// as plain text or put in by Replace, takes the formatting where it lands, as Word's Keep Text Only
// does, but a tab it brings is an ordinary tab.
import { Fragment, Slice, type Mark, type Node } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';

/** The marks for new text `text` given `marks`: without a positional tab when it holds a tab. */
export function insertedTextMarks(marks: readonly Mark[], text: string): readonly Mark[] {
	if (!text.includes('\t')) return marks;
	return marks.flatMap((mark) => {
		const props = mark.type.name === 'runProperties' ? mark.attrs.props : undefined;
		if (!props?.positionalTab) return [mark];
		const { positionalTab: _tab, ...rest } = props as Record<string, unknown>;
		return Object.keys(rest).length ? [mark.type.create({ ...mark.attrs, props: rest })] : [];
	});
}

function withOrdinaryTabs(fragment: Fragment): Fragment {
	const children: Node[] = [];
	fragment.forEach((node) => {
		if (node.isText) children.push(node.mark(insertedTextMarks(node.marks, node.text!)));
		else children.push(node.isLeaf ? node : node.copy(withOrdinaryTabs(node.content)));
	});
	return Fragment.fromArray(children);
}

/** Tabs pasted or dropped as plain text are ordinary tabs, also beside a positional tab. */
export function positionalTabInputPlugin(): Plugin {
	return new Plugin({
		props: {
			transformPasted: (slice, _view, plain) =>
				plain ? new Slice(withOrdinaryTabs(slice.content), slice.openStart, slice.openEnd) : slice,
		},
	});
}
