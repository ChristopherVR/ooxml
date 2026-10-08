export type PaintField =
	| 'fillPattern'
	| 'fillBackgroundColor'
	| 'fillTransparency'
	| 'linePattern'
	| 'lineTransparency';
export type PaintInput =
	| HTMLInputElement
	| (HTMLElement & {
			value: string;
			disabled: boolean;
			options: { value: string; label: string; disabled?: boolean }[];
	  });
export type PaintDialog = HTMLElement & {
	open: boolean;
	show(): void;
	close(): void;
	updateComplete: Promise<boolean>;
};
export type PaintButton = HTMLElement & { disabled: boolean };
export const paintFields: readonly PaintField[] = [
	'fillPattern',
	'fillBackgroundColor',
	'fillTransparency',
	'linePattern',
	'lineTransparency',
];
const labels: Record<PaintField, string> = {
	fillPattern: 'Fill pattern',
	fillBackgroundColor: 'Fill background color',
	fillTransparency: 'Fill transparency (%)',
	linePattern: 'Line pattern',
	lineTransparency: 'Line transparency (%)',
};

/** Typed presentation only; the core defines native pattern and transparency semantics. */
export function createPaintDialog(doc: Document) {
	const dialog = doc.createElement('office-ui-dialog') as PaintDialog;
	dialog.className = 'paint-properties-dialog';
	dialog.setAttribute('heading', 'Fill & Line');
	const fields = new Map<PaintField, PaintInput>();
	for (const [heading, subset] of [
		['Fill', paintFields.slice(0, 3)],
		['Line', paintFields.slice(3)],
	] as const) {
		const group = doc.createElement('fieldset');
		const legend = doc.createElement('legend');
		legend.textContent = heading;
		group.append(legend);
		for (const field of subset) {
			const label = doc.createElement('label');
			label.textContent = labels[field];
			let input: PaintInput;
			if (field.endsWith('Pattern')) {
				input = doc.createElement('office-ui-select') as Exclude<PaintInput, HTMLInputElement>;
				input.options = [];
			} else {
				input = doc.createElement('input');
				input.type = field === 'fillBackgroundColor' ? 'text' : 'number';
				if (field === 'fillBackgroundColor') {
					input.placeholder = '#RRGGBB';
					input.maxLength = 7;
				} else {
					input.min = '0';
					input.max = '100';
					input.step = '0.5';
				}
			}
			input.dataset.paintField = field;
			input.setAttribute('aria-label', labels[field]);
			label.append(input);
			group.append(label);
			fields.set(field, input);
		}
		dialog.append(group);
	}
	const hint = doc.createElement('p');
	hint.className = 'paint-properties-hint';
	hint.textContent =
		'Only changed fields are applied. Fill transparency changes foreground and background together. Source protection and formulas can refuse edits. Gradient transparency requires replacing the paint first.';
	const error = doc.createElement('p');
	error.setAttribute('role', 'alert');
	error.dataset.paintError = '';
	const buttons = ['Apply', 'Cancel'].map((label) => {
		const button = doc.createElement('office-ui-button') as PaintButton;
		button.slot = 'footer';
		button.setAttribute('label', label);
		button.setAttribute('command', `paint-${label.toLowerCase()}`);
		return button;
	});
	dialog.append(hint, error, ...buttons);
	return { dialog, fields, error, apply: buttons[0]!, cancel: buttons[1]! };
}
