import type { LinkedDataCard } from 'ooxml-core/xlsx';
import { sheetLocked, target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { field, text, textArea } from './fields';
import { showDialog } from './frame';

export async function editPictureAltText(
	ctx: EditorContext,
	current: NonNullable<ReturnType<typeof target>>,
	initial: string,
): Promise<void> {
	const input = textArea(ctx, initial, 4);
	input.maxLength = 1024;
	const alt = await showDialog(ctx, {
		name: 'in-cell-alt-text',
		heading: 'Alt Text',
		body: field(ctx, 'Alt Text', input),
		submit: () => input.value,
	});
	if (
		alt !== undefined &&
		!ctx.isCollaborating?.() &&
		!ctx.readOnly() &&
		target(ctx)?.ws === current.ws &&
		!sheetLocked(ctx)
	)
		current.session.setCellPictureAltText(current.sheet, current.active, alt);
}

export async function showLinkedData(ctx: EditorContext, card: LinkedDataCard): Promise<void> {
	const doc = ctx.host.ownerDocument;
	const fields = doc.createElement('dl');
	for (const { name, value } of card.fields) {
		const label = doc.createElement('dt');
		label.textContent = name;
		const detail = doc.createElement('dd');
		detail.textContent = value;
		fields.append(label, detail);
	}
	const name = doc.createElement('strong');
	name.textContent = card.text;
	await showDialog(ctx, {
		name: 'linked-data-card',
		heading: 'Linked Data Type',
		body: [
			name,
			fields,
			text(ctx, 'Refresh is unavailable. This workbook shows cached linked data.'),
		],
		okLabel: null,
	});
}
