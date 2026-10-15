import {
	SLICER_STYLE_NAMES,
	pixelsToEmu,
	resizedFrameAnchor,
	slicerStyleLabel,
	slicerStylePreview,
	type SlicerPatch,
	anchorToPixelBox,
} from 'ooxml-core/xlsx';
import {
	activeSlicer,
	editing,
	icon,
	slicerSourceAt,
	target,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';

const PIXELS_PER_INCH = 96;
const inches = (px: number): string => (px / PIXELS_PER_INCH).toFixed(2);

function update(ctx: EditorContext, patch: SlicerPatch): void {
	const t = target(ctx);
	const found = activeSlicer(ctx);
	if (!t || !found) return;
	try {
		t.session.updateSlicer(found.slicer.name, patch);
	} catch (error) {
		ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');
	}
}

function resize(
	ctx: EditorContext,
	size: { width?: number | undefined; height?: number | undefined },
): void {
	const t = target(ctx);
	const found = activeSlicer(ctx);
	if (!t || !found) return;
	t.session.setDrawingAnchor(
		t.sheet,
		found.index,
		resizedFrameAnchor(t.ws, found.frame.anchor, size),
	);
}

const inchValue = (raw: unknown): number | undefined => {
	const value = Number(String(raw ?? '').replace(',', '.'));
	return Number.isFinite(value) && value > 0 ? value * PIXELS_PER_INCH : undefined;
};

/** Slicer sizes offered by the Size controls, in inches (typed values are accepted too). */
export const SLICER_SIZE_OPTIONS = ['0.50', '1.00', '1.50', '2.00', '2.50', '3.00', '4.00', '5.00'];

export const slicerStyleGalleryItems = (ctx: EditorContext) =>
	SLICER_STYLE_NAMES.map((name) => {
		const { family, index } = slicerStyleLabel(name);
		return {
			id: name,
			label: ctx.t('{family} {n}', { family: ctx.t(family), n: index }),
			preview: slicerStylePreview(
				name,
				ctx.workbook()?.theme ?? { colors: [], majorFont: '', minorFont: '' },
			),
		};
	});

export function slicerCommands(): Command[] {
	const slicerCmd = (command: Omit<Command, 'editing'>): Command =>
		editing({
			...command,
			lock: 'objects',
			enabled: (ctx) => !!activeSlicer(ctx) && (command.enabled?.(ctx) ?? true),
		});
	return [
		editing({
			id: 'insert.slicer',
			label: 'Slicer',
			icon: icon('slicer'),
			lock: 'objects',
			enabled: (ctx) => !!slicerSourceAt(ctx),
			run: (ctx) => void ctx.dialogs.open('insert-slicer'),
		}),
		slicerCmd({
			id: 'slicer.settings',
			label: 'Slicer Settings',
			icon: icon('slicer'),
			run: (ctx) => void ctx.dialogs.open('slicer-settings'),
		}),
		slicerCmd({
			id: 'slicer.style',
			label: 'Slicer Styles',
			icon: icon('slicer'),
			run: (ctx, arg) => update(ctx, { style: String(arg) }),
		}),
		slicerCmd({
			id: 'slicer.columns',
			label: 'Columns',
			icon: icon('slicer'),
			value: (ctx) => String(activeSlicer(ctx)?.slicer.columnCount ?? 1),
			run: (ctx, arg) => update(ctx, { columnCount: Math.round(Number(arg)) }),
		}),
		slicerCmd({
			id: 'slicer.button-height',
			label: 'Button Height',
			icon: icon('slicer'),
			value: (ctx) => {
				const emu = activeSlicer(ctx)?.slicer.rowHeight;
				return inches(emu === undefined ? 26.44 : emu / 9525);
			},
			run: (ctx, arg) => {
				const px = inchValue(arg);
				if (px !== undefined) update(ctx, { rowHeight: pixelsToEmu(px) });
			},
		}),
		slicerCmd({
			id: 'slicer.height',
			label: 'Height',
			icon: icon('slicer'),
			value: (ctx) => {
				const t = target(ctx);
				const found = activeSlicer(ctx);
				return t && found ? inches(anchorToPixelBox(t.ws, found.frame.anchor).h) : '';
			},
			run: (ctx, arg) => resize(ctx, { height: inchValue(arg) }),
		}),
		slicerCmd({
			id: 'slicer.width',
			label: 'Width',
			icon: icon('slicer'),
			value: (ctx) => {
				const t = target(ctx);
				const found = activeSlicer(ctx);
				return t && found ? inches(anchorToPixelBox(t.ws, found.frame.anchor).w) : '';
			},
			run: (ctx, arg) => resize(ctx, { width: inchValue(arg) }),
		}),
		slicerCmd({
			id: 'slicer.delete',
			label: 'Delete Slicer',
			icon: icon('clear'),
			run: (ctx) => {
				const t = target(ctx);
				const found = activeSlicer(ctx);
				if (t && found) t.session.deleteSlicer(found.slicer.name);
			},
		}),
	];
}
