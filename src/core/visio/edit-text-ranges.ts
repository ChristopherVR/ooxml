import type { VisioPackage } from './package';
import type { VisioTextRangesEdit } from './edit-text-range-commands';
import { children, attribute } from './sheet';
import { fail } from './package-common';
import { localTextTarget } from './edit-text-target';
import { effectiveFormattingRows } from './edit-style-admission';
import { assertFormattingText } from './edit-formatting-rows';
import { assertFormattingDependencies } from './edit-formatting-scope';
import { textRangeNodes } from './edit-text-range-plan';

export async function replaceScopedTextRanges(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioTextRangesEdit,
	check: () => void,
): Promise<boolean> {
	const shape = localTextTarget(roots.get(edit.pageId)!, document, edit.shapeId, check);
	const texts = children(shape, 'Text');
	if (texts.length !== 1)
		fail('UNSUPPORTED_TEXT_RANGE', 'Range editing requires one existing local Text element.');
	const text = texts[0]!;
	// Plain source keeps the existing writer's admission: unrelated stored formatting rows
	// need no rich-marker provenance proof when there are no markers to interpret.
	if (Array.from(text.childNodes).some((node) => ![3, 4].includes(node.nodeType))) {
		const characters = effectiveFormattingRows(shape, document, 'Character'),
			paragraphs = effectiveFormattingRows(shape, document, 'Paragraph');
		assertFormattingText(shape, characters, paragraphs);
	}
	let preceding = '';
	for (const node of Array.from(text.childNodes)) {
		if (node.nodeType === 3 || node.nodeType === 4) preceding += node.nodeValue ?? '';
		else {
			const marker = node as Element;
			if (
				(marker.localName === 'pp' && preceding && !preceding.endsWith('\n')) ||
				(marker.localName === 'tp' && attribute(marker, 'IX') !== '0')
			)
				fail(
					'UNSUPPORTED_TEXT_RANGE',
					'Noncanonical paragraph or tab marker placement cannot be edited.',
				);
		}
	}
	const plan = textRangeNodes(text, edit);
	if (!plan.changed) return false;
	await assertFormattingDependencies(
		pkg,
		pagePaths,
		roots,
		edit.pageId,
		shape,
		new Map([['TheText', text]]),
		check,
	);
	check();
	while (text.firstChild) text.removeChild(text.firstChild);
	for (const node of plan.nodes) text.appendChild(node);
	return true;
}
