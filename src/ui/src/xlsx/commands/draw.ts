// Excel's Draw tab: Select, Lasso Select and Eraser, the shared pen strip (each pen's colour and
// thickness in its own menu) with Add Pen, Draw with Touch, Ruler, Ink to Shape, Ink to Math and
// Ink Replay. The state is the
// shared `DrawState`; drawing over the cells is the grid's ink layer (`grid/ink-layer.ts`).
import type { InkPen } from 'ooxml-core/drawingml';
import {
	editing,
	icon,
	unavailableReason,
	viewing,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { DrawState, inkPensLabels } from '../../ink/draw-state';
import type { InkPenChangeDetail, InkPenSelectDetail, OfficeUiInkPens } from '../../ink/ink-pens';
import { inkMathRecognizer } from '../../ink/math-recognizer';
import { registerInkElements } from '../../ink/register';

const states = new WeakMap<EditorContext, DrawState>();

/** The Draw tab's state of an editor, which drives its grid's ink layer. */
export function drawState(ctx: EditorContext): DrawState {
	let state = states.get(ctx);
	if (!state) {
		const created = new DrawState();
		created.onChange(() => {
			syncInkTool(ctx, created);
			ctx.requestRender();
		});
		states.set(ctx, created);
		state = created;
	}
	return state;
}

function syncInkTool(ctx: EditorContext, state: DrawState): void {
	const tool = state.canvasTool;
	ctx
		.grid()
		?.ink?.()
		.setTool(
			tool
				? { tool, color: state.pen.color, width: state.pen.width, touch: state.snapshot.touch }
				: undefined,
		);
}

const hasInk = (ctx: EditorContext): boolean =>
	Boolean(
		ctx.workbook()?.sheets[ctx.activeSheet()]?.drawings.some((drawing) => drawing.kind === 'ink'),
	);

const TOOL_DESCRIPTIONS = {
	select: 'Select objects and text instead of drawing.',
	lasso: 'Draw around ink to select it.',
	eraser: 'Erase the ink strokes you touch.',
} as const;

function tool(
	id: string,
	label: string,
	glyph: Parameters<typeof icon>[0],
	value: 'select' | 'lasso' | 'eraser',
): Command {
	const command = {
		id,
		label,
		description: TOOL_DESCRIPTIONS[value],
		icon: icon(glyph),
		lock: 'objects',
		unavailable: 'An editable sheet that allows objects is required.',
		checked: (ctx: EditorContext) => drawState(ctx).snapshot.tool === value,
		run: (ctx: EditorContext) => drawState(ctx).setTool(value),
	};
	return value === 'select' ? viewing(command) : editing(command);
}

/** Every Draw tab command. */
export function drawCommands(): Command[] {
	return [
		tool('draw.select', 'Select', 'drawSelect', 'select'),
		tool('draw.lasso', 'Lasso Select', 'lasso', 'lasso'),
		tool('draw.eraser', 'Eraser', 'eraser', 'eraser'),
		editing({
			id: 'draw.pens',
			label: 'Pens',
			icon: icon('penColor'),
			lock: 'objects',
			unavailable: 'An editable sheet that allows objects is required.',
			run: (ctx, arg) => {
				const [kind, index] = String(arg).split(':');
				const at = Number(index);
				if (!Number.isInteger(at)) return;
				if (kind === 'remove') drawState(ctx).removePen(at);
				else drawState(ctx).selectPen(at);
			},
		}),
		editing({
			id: 'draw.add-pen',
			label: 'Add Pen',
			icon: icon('addPen'),
			lock: 'objects',
			unavailable: (ctx) =>
				drawState(ctx).snapshot.pens.length >= 12
					? 'The pen gallery is full.'
					: 'An editable sheet that allows objects is required.',
			enabled: (ctx) => drawState(ctx).snapshot.pens.length < 12,
			run: (ctx, arg) => {
				const kind = ['pen', 'pencil', 'highlighter'].find((item) => item === arg) as
					| InkPen['kind']
					| undefined;
				drawState(ctx).addPen(kind ?? 'pen');
			},
		}),
		editing({
			id: 'draw.pen',
			label: 'Pens',
			lock: 'objects',
			unavailable: 'An editable sheet that allows objects is required.',
			run: (ctx, arg) => {
				const change = arg as { color?: string; width?: number } | undefined;
				if (change) drawState(ctx).updatePen(change);
			},
		}),
		viewing({
			id: 'draw.ruler',
			label: 'Ruler',
			icon: icon('ruler'),
			unavailable: 'The ruler is not available in this editor.',
			enabled: () => false,
			run: () => undefined,
		}),
		editing({
			id: 'draw.touch',
			description: 'Draw with your finger as well as a pen or mouse.',
			label: 'Draw with Touch',
			icon: icon('drawTouch'),
			lock: 'objects',
			unavailable: 'An editable sheet that allows objects is required.',
			checked: (ctx) => drawState(ctx).snapshot.touch,
			run: (ctx) => {
				const state = drawState(ctx);
				state.setTouch(!state.snapshot.touch);
			},
		}),
		editing({
			id: 'draw.ink-to-shape',
			description: 'Turn ink you drew into a shape.',
			label: 'Ink to Shape',
			icon: icon('inkToShape'),
			lock: 'objects',
			unavailable: 'Select ink with Lasso Select or select an ink drawing.',
			enabled: (ctx) => ctx.grid()?.ink?.().hasSelection() ?? false,
			run: async (ctx) => {
				const done = await ctx.grid()?.ink?.().inkToShape();
				if (done === false)
					ctx.toast(ctx.t('Ink to Shape did not recognize a shape in the selected ink.'), 'info');
			},
		}),
		editing({
			id: 'draw.ink-to-math',
			description: 'Turn handwritten math into an equation.',
			label: 'Ink to Math',
			icon: icon('inkToMath'),
			lock: 'objects',
			unavailable: () =>
				inkMathRecognizer()
					? 'An editable sheet that allows objects is required.'
					: 'Ink to Math needs a handwriting recognizer, which this editor does not include.',
			enabled: () => Boolean(inkMathRecognizer()),
			run: async (ctx) => {
				const recognizer = inkMathRecognizer();
				if (!recognizer) return;
				const { openInkEquation } = await import('./equation');
				await openInkEquation(ctx, recognizer);
			},
		}),
		viewing({
			id: 'draw.replay',
			description: 'Watch the ink being drawn again, in the order it was drawn.',
			label: 'Ink Replay',
			icon: icon('inkReplay'),
			unavailable: 'There is no ink on this sheet to replay.',
			enabled: (ctx) => hasInk(ctx) || (ctx.grid()?.ink?.().replaying() ?? false),
			checked: (ctx) => ctx.grid()?.ink?.().replaying() ?? false,
			run: (ctx) => {
				drawState(ctx).setTool('select');
				ctx.grid()?.ink?.().replay();
			},
		}),
	];
}

/** The Draw tab's pen strip: the shared `office-ui-ink-pens`, run through the Draw commands. */
export function renderPenStrip(
	ctx: EditorContext,
	update: (refresh: () => void) => void,
): HTMLElement {
	const doc = ctx.root.ownerDocument;
	const strip = doc.createElement('office-ui-ink-pens') as OfficeUiInkPens;
	strip.addEventListener('office-ink-pen-select', (event) => {
		void ctx.commands.run(
			'draw.pens',
			`pen:${(event as CustomEvent<InkPenSelectDetail>).detail.index}`,
		);
	});
	strip.addEventListener('office-ink-pen-remove', (event) => {
		void ctx.commands.run(
			'draw.pens',
			`remove:${(event as CustomEvent<{ index: number }>).detail.index}`,
		);
	});
	strip.addEventListener('office-ink-pen-change', (event) => {
		const { color, width } = (event as CustomEvent<InkPenChangeDetail>).detail;
		void ctx.commands.run('draw.pen', { ...(color ? { color } : {}), ...(width ? { width } : {}) });
	});
	// The strip's element loads when the Draw tab is first shown, as Word's does.
	let defined = false;
	const define = () => {
		const ribbon = strip.closest<HTMLElement & { selected?: string }>('office-ui-ribbon');
		if (defined || !strip.isConnected || ribbon?.selected !== 'draw') return;
		defined = true;
		shown?.disconnect();
		void registerInkElements(doc.defaultView?.customElements);
	};
	const Observer = doc.defaultView?.IntersectionObserver;
	const shown = Observer ? new Observer(define) : undefined;
	shown?.observe(strip);
	update(() => {
		define();
		const snapshot = drawState(ctx).snapshot;
		const command = ctx.commands.get('draw.pens');
		const enabled = ctx.commands.isEnabled('draw.pens');
		strip.labels = inkPensLabels((text) => ctx.t(text), ctx.locale());
		strip.pens = snapshot.pens;
		strip.selected = snapshot.tool === 'pen' ? snapshot.pen : -1;
		strip.disabled = !enabled;
		strip.disabledReason = !enabled && command ? (unavailableReason(ctx, command) ?? '') : '';
	});
	return strip;
}
