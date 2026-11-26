import type { NodeSpec } from 'prosemirror-model';

/** A comment at a caret occupies no document text. */
export const commentPointNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: false,
	marks: '',
	attrs: { id: { default: '' } },
	leafText: () => '',
	parseDOM: [
		{
			tag: 'span[data-docx-comment-point]',
			getAttrs: (element) => ({
				id: element.getAttribute('data-docx-comment-point') ?? '',
			}),
		},
	],
	toDOM: (node) => [
		'span',
		{
			'data-docx-comment-point': node.attrs.id,
			'aria-hidden': 'true',
			contenteditable: 'false',
		},
	],
};
