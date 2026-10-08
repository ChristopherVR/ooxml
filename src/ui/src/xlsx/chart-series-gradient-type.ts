import type { DrawingGradientGeometryType } from 'ooxml-core/diagram';
import type { DrawingFill } from 'ooxml-core/drawingml';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select } from './dialogs/fields';

/** A native type field that keeps unsupported imported types visible until explicitly changed. */
export function createGradientTypeField(
	ctx: EditorContext,
	onPick: (type: DrawingGradientGeometryType) => void,
) {
	const input = select(ctx, [
		['linear', 'Linear'],
		['circle', 'Radial'],
		['rect', 'Rectangular'],
		['shape', 'Path'],
	]);
	input.options[1]!.disabled = true;
	input.options[3]!.disabled = true;
	const other = el(ctx, 'option');
	other.value = 'unsupported';
	other.disabled = true;
	input.append(other);
	const element = field(ctx, 'Type', input);
	input.addEventListener('change', () => {
		if (
			!input.disabled &&
			!input.selectedOptions[0]?.disabled &&
			['linear', 'rect', 'circle', 'shape'].includes(input.value)
		)
			onPick(input.value as DrawingGradientGeometryType);
	});
	return {
		element,
		refresh(
			fill: Extract<DrawingFill, { kind: 'gradient' }>,
			disabled: boolean,
			rectangularMarks = false,
		) {
			input.options[1]!.disabled = input.options[3]!.disabled = !rectangularMarks;
			element.querySelector('span')!.textContent = ctx.t('Type');
			input.setAttribute('aria-label', ctx.t('Type'));
			const linearLabel = ctx.t('Linear gradient');
			input.options[0]!.textContent = linearLabel === 'Linear gradient' ? 'Linear' : linearLabel;
			input.options[1]!.textContent = ctx.t('Radial');
			input.options[2]!.textContent = ctx.t('Rectangular');
			input.options[3]!.textContent = ctx.t('Path');
			other.textContent = ctx.t('Imported gradient type');
			input.value = !fill.path
				? 'linear'
				: ['rect', 'circle', 'shape'].includes(fill.path)
					? fill.path
					: 'unsupported';
			other.hidden = input.value !== 'unsupported';
			input.disabled = disabled;
		},
	};
}
