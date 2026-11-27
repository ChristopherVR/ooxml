import { pictureAnchorAt } from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { createEquationFields } from '../../math/equation-fields';
import { equationPreview } from '../../math/equation-render';
import type { InkMathRecognizer } from '../../ink/math-recognizer';
import { showDialog } from '../dialogs/frame';

export async function openEquation(ctx: EditorContext, edit: boolean) {
	const t = target(ctx);
	if (!t) return;
	const index = edit ? ctx.selection.get().drawing : undefined;
	const drawing = index === undefined ? undefined : t.ws.drawings[index];
	const equations =
		drawing?.kind === 'unsupported'
			? (drawing.shape?.equationParagraphs?.flat().filter((item) => item.kind === 'equation') ?? [])
			: [];
	if (edit && !equations.some((item) => item.omml)) return;
	const doc = ctx.host.ownerDocument;
	const body = doc.createElement('div');
	const mount = doc.createElement('div');
	const select = doc.createElement('select');
	const label = doc.createElement('label');
	label.append(ctx.t('Equation'), select);
	if (equations.length > 1) {
		for (const [i, equation] of equations.entries()) {
			const option = doc.createElement('option');
			option.value = String(i);
			option.disabled = !equation.omml;
			option.textContent = `${i + 1}. ${equation.fallback}`;
			select.append(option);
		}
		select.value = String(equations.findIndex((item) => item.omml));
		body.append(label);
	}
	body.append(mount);
	let confirm: HTMLButtonElement | null = null;
	const createFields = () => {
		const current = equations[Number(select.value)];
		const latex = current?.omml ? equationPreview(current.omml)?.latex || current.fallback : '';
		const fields = createEquationFields(doc, {
			labels: {
				equation: ctx.t('Equation'),
				source: ctx.t('LaTeX equation'),
				display: '',
				preview: ctx.t('Preview'),
				empty: ctx.t('Type a LaTeX equation to see a preview.'),
				invalid: ctx.t('This LaTeX could not be converted to an equation.'),
			},
			latex,
			display: true,
			onChange: (value) => {
				if (confirm) confirm.disabled = !value;
			},
			accept: () => confirm?.click(),
		});
		mount.replaceChildren(fields.element);
		return { fields, latex };
	};
	let current = createFields();
	select.addEventListener('change', () => {
		current = createFields();
	});
	await showDialog(
		ctx,
		{
			name: 'equation',
			heading: edit ? 'Edit Equation' : 'Equation',
			body,
			opened: () => current.fields.source.select(),
			submit: () => {
				const value = current.fields.draft;
				const live = target(ctx);
				const command = edit ? 'drawing.edit-equation' : 'insert.equation';
				if (
					!value ||
					live?.session !== t.session ||
					live.ws !== t.ws ||
					!ctx.commands.isEnabled(command)
				)
					return;
				const equation = { omml: value.omml, fallback: value.latex };
				if (index !== undefined) {
					if (live.ws.drawings[index] !== drawing || ctx.selection.get().drawing !== index) return;
					return (
						value.latex === current.latex ||
						t.session.setShapeEquation(live.sheet, index, Number(select.value), equation) ||
						undefined
					);
				}
				const added = t.session.addShape(live.sheet, pictureAnchorAt(t.active, 240, 72), {
					kind: 'equation',
					name: ctx.t('Equation'),
					equation,
				});
				ctx.selection.set({ drawing: added });
				return true;
			},
		},
		({ element }) => {
			confirm = element.querySelector('button.xve-btn-primary');
			if (confirm) confirm.disabled = !current.fields.draft;
		},
	);
}

/**
 * Draw > Ink to Math: the shared Ink Equation dialog with the host's recognizer, as Excel opens its
 * Math Input Control; the equation goes in a new equation text box at the active cell.
 */
export async function openInkEquation(ctx: EditorContext, recognizer: InkMathRecognizer) {
	const t = target(ctx);
	if (!t) return;
	const { openInkEquationDialog } = await import('../../ink/ink-equation-dialog');
	openInkEquationDialog({
		host: ctx.root,
		recognizer,
		locale: ctx.locale(),
		heading: 'Ink to Math',
		insert: (draft) => {
			const live = target(ctx);
			if (live?.session !== t.session || !ctx.commands.isEnabled('draw.ink-to-math')) return false;
			const added = live.session.addShape(live.sheet, pictureAnchorAt(live.active, 240, 72), {
				kind: 'equation',
				name: ctx.t('Equation'),
				equation: { omml: draft.omml, fallback: draft.latex },
			});
			ctx.selection.set({ drawing: added });
			return true;
		},
	});
}
