// Word's Draw tab, grouped as Microsoft 365's: Drawing Tools (Select, Lasso Select, Eraser, the
// shared pen strip with each pen's colour and thickness, Add Pen, Draw with Touch), Stencils,
// Convert (Ink to Shape, Ink to Math) and Replay. The state is
// the shared `DrawState`; drawing on the pages is `draw-controller.ts`.
import { inkPensLabels, type DrawState } from '../ink/draw-state';
import type { InkPenChangeDetail, InkPenSelectDetail, OfficeUiInkPens } from '../ink/ink-pens';
import { emit } from './events';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import type { DrawKey } from './locales/draw';

/** A Draw tab string: the editor's own or one of the tab's (merged when the tab loads). */
export type DrawText = LocalizationKey | DrawKey;
import type { RibbonAction } from './ribbon-action';
import { group, menuSelect, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';

/** Why a Draw command cannot run now, or undefined when it can. */
export interface DrawAvailability {
	/** Drawing on the pages (tools, pens, Draw with Touch). */
	draw?: DrawText | undefined;
	inkToShape?: DrawText | undefined;
	inkToMath?: DrawText | undefined;
	replay?: DrawText | undefined;
	/** Ink Replay is running (its button shows pressed). */
	replaying?: boolean;
}

const marker = (element: HTMLElement, key: string) => {
	element.dataset.draw = key;
	return element;
};

/** The pen strip's element loads with the Draw tab: an editor that never draws never downloads it. */
export function defineDrawElements(): void {
	void import('../ink/ink-pens').then(({ defineInkPens }) => defineInkPens());
}

export function buildDrawPanel(panel: HTMLElement): void {
	const pens = document.createElement('office-ui-ink-pens') as OfficeUiInkPens;
	pens.dataset.draw = 'pens';
	pens.addEventListener('office-ink-pen-select', (event) =>
		emit(pens, 'ribbon-action', {
			type: 'draw',
			key: 'pen',
			index: (event as CustomEvent<InkPenSelectDetail>).detail.index,
		}),
	);
	pens.addEventListener('office-ink-pen-change', (event) => {
		const { color, width } = (event as CustomEvent<InkPenChangeDetail>).detail;
		if (color) emit(pens, 'ribbon-action', { type: 'draw', key: 'penColor', value: color });
		if (width) emit(pens, 'ribbon-action', { type: 'draw', key: 'thickness', value: width });
	});
	pens.addEventListener('office-ink-pen-remove', (event) =>
		emit(pens, 'ribbon-action', {
			type: 'draw',
			key: 'removePen',
			index: (event as CustomEvent<{ index: number }>).detail.index,
		}),
	);
	const addPen = menuSelect(
		'Add Pen',
		'addPen',
		[
			['pen', 'Pen'],
			['pencil', 'Pencil'],
			['highlighter', 'Highlighter'],
		],
		(value) => ({ type: 'draw', key: 'addPen', value: value as 'pen' | 'pencil' | 'highlighter' }),
		{ momentary: true },
	);
	addPen.dataset.draw = 'addPen';
	const large = (label: string, icon: RibbonIcon, action: RibbonAction | null, key: string) =>
		marker(tool(label, icon, action, { large: true }), key);
	panel.append(
		group(
			'Drawing Tools',
			large('Select', 'select', { type: 'draw', key: 'tool', value: 'select' }, 'select'),
			large('Lasso Select', 'lasso', { type: 'draw', key: 'tool', value: 'lasso' }, 'lasso'),
			large('Eraser', 'eraser', { type: 'draw', key: 'tool', value: 'eraser' }, 'eraser'),
			pens,
			addPen,
			large('Draw with Touch', 'drawTouch', { type: 'draw', key: 'touch' }, 'touch'),
		),
		group('Stencils', large('Ruler', 'ruler', null, 'ruler')),
		group(
			'Convert',
			large('Ink to Shape', 'inkToShape', { type: 'draw', key: 'inkToShape' }, 'inkToShape'),
			large('Ink to Math', 'inkToMath', { type: 'draw', key: 'inkToMath' }, 'inkToMath'),
		),
		group('Replay', large('Ink Replay', 'inkReplay', { type: 'draw', key: 'replay' }, 'replay')),
	);
}

/** What each Draw command does, shown under its name in the tooltip. */
const DESCRIPTIONS: Record<string, DrawText> = {
	select: 'Select objects and text instead of drawing.',
	lasso: 'Draw around ink to select it.',
	eraser: 'Erase the ink strokes you touch.',
	touch: 'Draw with your finger as well as a pen or mouse.',
	inkToShape: 'Turn ink you drew into a shape.',
	inkToMath: 'Turn handwritten math into an equation.',
	replay: 'Watch the ink being drawn again, in the order it was drawn.',
};

function describe(
	control: HTMLButtonElement | HTMLSelectElement,
	locale: EditorLocale,
	reason: DrawText | undefined,
): void {
	control.disabled = Boolean(reason);
	const label =
		control.dataset.localearialabel ?? control.getAttribute('aria-label') ?? control.title;
	if (reason) {
		const text = translate(locale, reason as LocalizationKey);
		control.title = `${translate(locale, label as LocalizationKey)}: ${text}`;
		control.setAttribute('aria-description', text);
	} else {
		const description = DESCRIPTIONS[control.dataset.draw ?? ''];
		control.title = translate(locale, label as LocalizationKey);
		if (description) {
			control.title += `\n${translate(locale, description as LocalizationKey)}`;
			control.setAttribute('aria-description', translate(locale, description as LocalizationKey));
		} else control.removeAttribute('aria-description');
	}
}

/** Pressed tools, the pen gallery, the pen's thickness and each command's availability. */
export function syncDrawRibbon(
	toolbar: HTMLElement,
	state: DrawState,
	availability: DrawAvailability,
	locale: EditorLocale,
): void {
	const panel = toolbar.querySelector<HTMLElement>('[data-ribbon-tab="draw"]');
	if (!panel) return;
	const t = (text: string) => translate(locale, text as LocalizationKey);
	const snapshot = state.snapshot;
	const button = (key: string) =>
		panel.querySelector<HTMLButtonElement>(`button[data-draw="${key}"]`);
	for (const key of ['select', 'lasso', 'eraser'] as const) {
		const control = button(key);
		if (!control) continue;
		control.setAttribute('aria-pressed', String(snapshot.tool === key));
		describe(control, locale, key === 'select' ? undefined : availability.draw);
	}
	const touch = button('touch');
	if (touch) {
		touch.setAttribute('aria-pressed', String(snapshot.touch));
		describe(touch, locale, availability.draw);
	}
	for (const [key, reason] of [
		['inkToShape', availability.inkToShape],
		['inkToMath', availability.inkToMath],
		['replay', availability.replay],
	] as const) {
		const control = button(key);
		if (control) describe(control, locale, reason);
	}
	button('replay')?.setAttribute('aria-pressed', String(availability.replaying ?? false));
	const pens = panel.querySelector<OfficeUiInkPens>('office-ui-ink-pens[data-draw="pens"]');
	if (pens) {
		pens.labels = inkPensLabels(t);
		pens.pens = snapshot.pens;
		pens.selected = snapshot.tool === 'pen' ? snapshot.pen : -1;
		pens.disabled = Boolean(availability.draw);
		pens.disabledReason = availability.draw ? t(availability.draw) : '';
	}
	const ruler = button('ruler');
	if (ruler) describe(ruler, locale, 'The ruler is not available in this editor.');
	const addPen = panel.querySelector<HTMLSelectElement>('[data-draw="addPen"] select');
	if (addPen)
		describe(
			addPen,
			locale,
			availability.draw ?? (snapshot.pens.length >= 12 ? 'The pen gallery is full.' : undefined),
		);
}
