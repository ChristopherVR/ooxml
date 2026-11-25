import {
	canGroupDrawings,
	canRotateDrawing,
	canUngroupDrawing,
	editing,
	icon,
	target,
	viewing,
	type Command,
	type DrawingAlignment,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { selectedDrawings, setDrawingSelection } from '../drawing-selection';
import { toggleSelectionPane } from '../selection-pane';

export const ALIGNMENTS: ReadonlyArray<readonly [DrawingAlignment, string]> = [
	['left', 'Align Left'],
	['center', 'Align Center'],
	['right', 'Align Right'],
	['top', 'Align Top'],
	['middle', 'Align Middle'],
	['bottom', 'Align Bottom'],
];

const primary = (ctx: EditorContext) => selectedDrawings(ctx).at(-1);
const locked = 'objects';
const hasTransformableSelection = (ctx: EditorContext) => {
	const sheet = target(ctx)?.ws;
	return selectedDrawings(ctx).some((index) => canRotateDrawing(sheet?.drawings[index]));
};

export function arrangeCommands(): Command[] {
	return [
		...(
			[
				[1, 'forward', 'Bring Forward'],
				[-1, 'backward', 'Send Backward'],
			] as const
		).map(([direction, id, label]) =>
			editing({
				id: 'arrange.' + id,
				label,
				icon: icon(direction === 1 ? 'bringForward' : 'sendBackward'),
				lock: locked,
				enabled: (ctx) => {
					const index = primary(ctx);
					return (
						index !== undefined &&
						selectedDrawings(ctx).length === 1 &&
						!!target(ctx)?.ws.drawings[index + direction]
					);
				},
				run: (ctx) => {
					const t = target(ctx);
					const index = primary(ctx);
					if (t && index !== undefined)
						setDrawingSelection(ctx, [t.session.reorderDrawing(t.sheet, index, direction)]);
				},
			}),
		),
		viewing({
			id: 'arrange.selection-pane',
			label: 'Selection Pane',
			icon: icon('selectionPane'),
			run: (ctx) => toggleSelectionPane(ctx),
		}),
		...ALIGNMENTS.map(([alignment, label]) =>
			editing({
				id: 'arrange.align-' + alignment,
				label,
				lock: locked,
				enabled: (ctx) => {
					const sheet = target(ctx)?.ws;
					return (
						selectedDrawings(ctx).filter((index) => {
							const drawing = sheet?.drawings[index];
							return drawing && drawing.kind !== 'formControl';
						}).length >= 2
					);
				},
				run: (ctx) => {
					const t = target(ctx);
					if (t) t.session.alignDrawings(t.sheet, selectedDrawings(ctx), alignment);
				},
			}),
		),
		editing({
			id: 'arrange.group',
			label: 'Group',
			lock: locked,
			enabled: (ctx) => {
				const t = target(ctx);
				return !!t && canGroupDrawings(t.ws, selectedDrawings(ctx));
			},
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				const index = t.session.groupDrawings(t.sheet, selectedDrawings(ctx), ctx.t('Group'));
				if (index !== undefined) setDrawingSelection(ctx, [index]);
			},
		}),
		editing({
			id: 'arrange.ungroup',
			label: 'Ungroup',
			lock: locked,
			enabled: (ctx) => {
				const index = primary(ctx);
				return (
					index !== undefined &&
					selectedDrawings(ctx).length === 1 &&
					canUngroupDrawing(target(ctx)?.ws.drawings[index])
				);
			},
			run: (ctx) => {
				const t = target(ctx);
				const index = primary(ctx);
				if (t && index !== undefined)
					setDrawingSelection(ctx, t.session.ungroupDrawing(t.sheet, index));
			},
		}),
		...(
			[
				[90, 'right', 'Rotate Right 90°'],
				[-90, 'left', 'Rotate Left 90°'],
			] as const
		).map(([degrees, id, label]) =>
			editing({
				id: 'arrange.rotate-' + id,
				label,
				lock: locked,
				enabled: hasTransformableSelection,
				run: (ctx) => {
					const t = target(ctx);
					if (!t) return;
					t.session.batch(label, () => {
						for (const index of selectedDrawings(ctx))
							t.session.rotateDrawing(t.sheet, index, degrees);
					});
				},
			}),
		),
		...(
			[
				['horizontal', 'Flip Horizontal'],
				['vertical', 'Flip Vertical'],
			] as const
		).map(([axis, label]) =>
			editing({
				id: 'arrange.flip-' + axis,
				label,
				lock: locked,
				enabled: hasTransformableSelection,
				run: (ctx) => {
					const t = target(ctx);
					if (!t) return;
					t.session.batch(label, () => {
						for (const index of selectedDrawings(ctx)) t.session.flipDrawing(t.sheet, index, axis);
					});
				},
			}),
		),
	];
}
