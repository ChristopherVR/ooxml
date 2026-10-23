import type { GradientState } from 'ooxml-ui/pptx';
import {
	defaultGradientState,
	gradientStopColorCommitPatch,
	PATTERN_PRESET_OPTIONS,
	SHAPE_LINE_SECTIONS,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import type { ColorControlHandle, NumberFieldHandle } from '../controls';
import { makeButton, makeColorControl, makeNumberField } from '../controls';
import { createRecentColorsRow } from '../recent-colors-row';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
import { makeCheckboxField, makeRangeField, makeSelectField } from './controls-extra';
import { createInspectorFields } from './inspector-fields';
import { createShapeEffectsControls } from './shape-effects-controls';
import type { InspectorHandlers, InspectorState } from './types';

/** Preset used when a shape switches into pattern-fill mode with no prior one. */
const DEFAULT_PATTERN_PRESET = 'pct20';

export interface FillSection {
	el: HTMLElement;
	update(state: InspectorState): void;
}

/** Percent-formatted readout for the opacity sliders (0..1 -> "NN%"). */
function pct(value: number): string {
	return `${Math.round(value * 100)}%`;
}

/**
 * The Fill & Stroke section: flat colour + width (as before), plus opacity
 * sliders and a gradient-fill sub-panel (2-3 stop colour + angle), reusing
 * `pptx-viewer-shared`'s gradient-picker patch builders.
 */
export function createFillSection(
	doc: Document,
	t: Translator,
	section: (label: string) => HTMLElement,
	handlers: InspectorHandlers,
): FillSection {
	const el = section(t('pptx.shape.fillStroke'));
	// Scroll target for the "Format Shape" element context-menu command; see
	// `pptx-viewer-shared`'s `context-menu-inspector-anchor.ts`.
	el.setAttribute('data-pptx-inspector-section', 'fill-stroke');

	// A drag through the OS picker fires `input` on every shade it passes: those
	// show live with no undo step, and the settled colour (`change`, or `blur`
	// for a drag no change ended) is the one undo step and the one recent colour.
	const liveColor = (target: 'fill' | 'line', label: string, fallback: string) =>
		makeColorControl(
			doc,
			{
				label,
				onInput: (hex) => handlers.liveColor(target, 'input', hex),
				onCommit: (hex) => handlers.liveColor(target, 'change', hex),
				onBlur: (hex) => handlers.liveColor(target, 'blur', hex),
			},
			fallback,
		);
	const fill = liveColor('fill', t('pptx.inspector.fill'), '#4f86ff');
	const stroke = liveColor('line', t('pptx.inspector.line'), '#1e3a8a');

	// W3-G2: the deck's real "Theme Colors" grid, above the recent-colours
	// row. Clicking a theme swatch commits BOTH the resolved hex and the ref
	// (so the fill/stroke keeps following the theme after a later theme
	// change); every other commit path below clears the ref.
	const fillTheme = createThemeColorSwatchGrid(doc, t, (commit) =>
		handlers.setShapeFill(commit.hex, commit.ref),
	);
	const strokeTheme = createThemeColorSwatchGrid(doc, t, (commit) =>
		handlers.setShapeStroke(commit.hex, commit.ref),
	);

	// B6 (A1/A2): "Recent colours" rows under the fill and stroke pickers.
	// Clicking a swatch commits through the SAME handler the picker's own
	// `<input type="color">` uses (`setShapeFill`/`setShapeStroke`), which
	// folds a committed pick back into the deck's MRU list.
	const fillRecent = createRecentColorsRow(doc, t, handlers.setShapeFill);
	const strokeRecent = createRecentColorsRow(doc, t, handlers.setShapeStroke);

	// One labelled group per target, as React, Vue, Angular and Svelte lay it
	// out and as PowerPoint's Format Shape pane separates Fill from Line: the
	// picker, then that target's theme grid and recent colours. The two grids
	// used to be appended one after the other under a single Fill/Line row, so
	// nothing said which grid set the fill and which the line.
	const colorGroup = (label: string, control: HTMLElement, ...rest: HTMLElement[]): void => {
		const group = createEl(doc, 'div', 'pptxv-inspector-color-group');
		group.setAttribute('role', 'group');
		group.setAttribute('aria-label', label);
		const head = createEl(doc, 'div', 'pptxv-inspector-row');
		const title = createEl(doc, 'span', 'pptxv-inspector-row-label');
		title.textContent = label;
		head.append(title, control);
		group.append(head, ...rest);
		el.appendChild(group);
	};
	colorGroup(t('pptx.inspector.fill'), fill.el, fillTheme.el, fillRecent.el);
	colorGroup(t('pptx.inspector.line'), stroke.el, strokeTheme.el, strokeRecent.el);

	const fillOpacity = makeRangeField(doc, {
		label: t('pptx.strokeEffects.fillOpacity'),
		min: 0,
		max: 1,
		step: 0.01,
		formatValue: pct,
		onCommit: handlers.setFillOpacity,
	});
	el.appendChild(fillOpacity.el);
	// Line options (width, transparency, dash, compound, join, cap): the shared
	// descriptors, drawn by the generic renderer.
	const lineFields = createInspectorFields(doc, t, {
		sections: SHAPE_LINE_SECTIONS,
		onPatch: (patch) => handlers.setShapeStyle(patch),
	});
	el.appendChild(lineFields.el);

	// -- Gradient sub-panel -----------------------------------------------------
	const gradientToggle = makeCheckboxField(doc, {
		label: t('pptx.fill.gradient'),
		onChange(checked) {
			if (checked) {
				handlers.setGradientFill(lastGradient);
			} else {
				handlers.setShapeFill(fill.el.querySelector('input')?.value ?? '#4f86ff');
			}
		},
	});
	el.appendChild(gradientToggle.el);

	const gradientPanel = createEl(doc, 'div', 'pptxv-inspector-gradient');
	el.appendChild(gradientPanel);

	const typeSelectRow = createEl(doc, 'div', 'pptxv-inspector-row');
	gradientPanel.appendChild(typeSelectRow);
	const linearBtn = makeButton(doc, {
		label: t('pptx.gradient.linear'),
		text: t('pptx.gradient.linear'),
		onClick: () => handlers.setGradientFill({ ...lastGradient, type: 'linear' }),
	});
	const radialBtn = makeButton(doc, {
		label: t('pptx.gradient.radial'),
		text: t('pptx.gradient.radial'),
		onClick: () => handlers.setGradientFill({ ...lastGradient, type: 'radial' }),
	});
	typeSelectRow.append(linearBtn.btn, radialBtn.btn);

	const angleField = makeNumberField(doc, {
		label: t('pptx.gradient.angle'),
		min: 0,
		max: 360,
		onCommit: (angle) => handlers.setGradientFill({ ...lastGradient, angle }),
	});
	gradientPanel.appendChild(angleField.el);

	const stopsContainer = createEl(doc, 'div', 'pptxv-inspector-gradient-stops');
	gradientPanel.appendChild(stopsContainer);

	const addStopBtn = makeButton(doc, {
		label: t('pptx.gradient.addStop'),
		text: t('pptx.gradient.addStop'),
		onClick: () => {
			const stops = lastGradient.stops;
			const lastPos = stops[stops.length - 1]?.position ?? 100;
			const prevPos = stops[stops.length - 2]?.position ?? 0;
			handlers.addGradientStop('#ffffff', Math.round((lastPos + prevPos) / 2));
		},
	});
	gradientPanel.appendChild(addStopBtn.btn);

	// -- Pattern sub-panel --------------------------------------------------------
	// `a:pattFill`: the shape's own `fillColor` is the pattern foreground (the
	// existing Fill swatch above doubles as it, matching how the wire format
	// itself has no separate foreground field), `fillPatternBackgroundColor` is
	// the second colour, and `fillPatternPreset` picks one of the 56 presets.
	const patternToggle = makeCheckboxField(doc, {
		label: t('pptx.table.fillPattern'),
		onChange(checked) {
			if (checked) {
				handlers.setShapeStyle({
					fillMode: 'pattern',
					fillPatternPreset: lastPatternPreset,
					fillPatternBackgroundColor: lastPatternBackground,
				});
			} else {
				handlers.setShapeFill(fill.el.querySelector('input')?.value ?? '#4f86ff');
			}
		},
	});
	el.appendChild(patternToggle.el);

	const patternPanel = createEl(doc, 'div', 'pptxv-inspector-pattern');
	el.appendChild(patternPanel);

	const patternPreset = makeSelectField<string>(doc, {
		label: t('pptx.table.patternPreset'),
		options: PATTERN_PRESET_OPTIONS.map((opt) => ({ value: opt.value, label: t(opt.labelKey) })),
		onChange: (value) => {
			lastPatternPreset = value;
			handlers.setShapeStyle({ fillMode: 'pattern', fillPatternPreset: value });
		},
	});
	patternPanel.appendChild(patternPreset.el);

	const patternBackground = makeColorControl(
		doc,
		{
			label: t('pptx.table.patternBackground'),
			onInput: (hex) => {
				lastPatternBackground = hex;
				handlers.setShapeStyle({ fillMode: 'pattern', fillPatternBackgroundColor: hex });
			},
			onCommit: handlers.pushRecentColor,
		},
		'#ffffff',
	);
	patternPanel.appendChild(patternBackground.el);

	let lastPatternPreset = DEFAULT_PATTERN_PRESET;
	let lastPatternBackground = '#ffffff';

	let lastGradient: GradientState = defaultGradientState();
	let stopRows: Array<{
		color: ColorControlHandle;
		theme: ReturnType<typeof createThemeColorSwatchGrid>;
		position: NumberFieldHandle;
		remove: ReturnType<typeof makeButton>;
	}> = [];

	/**
	 * Rebuilds every stop row from scratch on each `update()` (stop count can
	 * change). Each stop gets its own "Theme Colors" grid under the native
	 * colour input, same "swatch commits hex + ref, native input clears it"
	 * contract as the fill/stroke pickers above.
	 */
	const rebuildStopRows = (
		stops: GradientState['stops'],
		themeColorMap: Record<string, string> | undefined,
		canShape: boolean,
	): void => {
		stopsContainer.replaceChildren();
		stopRows = stops.map((stop, index) => {
			const row = createEl(doc, 'div', 'pptxv-inspector-row');
			const color = makeColorControl(
				doc,
				{
					label: t('pptx.gradient.stops'),
					onInput: (hex) => handlers.updateGradientStop(index, { color: hex, colorRef: undefined }),
					onCommit: handlers.pushRecentColor,
				},
				stop.color,
			);
			const position = makeNumberField(doc, {
				label: t('pptx.gradient.position'),
				min: 0,
				max: 100,
				onCommit: (value) => handlers.updateGradientStop(index, { position: value }),
			});
			position.setValue(stop.position);
			const remove = makeButton(doc, {
				label: t('pptx.gradient.removeStop'),
				text: '✕',
				onClick: () => handlers.removeGradientStop(index),
			});
			remove.setDisabled(stops.length <= 2);
			row.append(color.el, position.el, remove.btn);
			stopsContainer.appendChild(row);
			const theme = createThemeColorSwatchGrid(doc, t, (commit) =>
				handlers.updateGradientStop(index, gradientStopColorCommitPatch(commit)),
			);
			theme.setThemeColorMap(themeColorMap);
			theme.setSelected(stop.colorRef, stop.color);
			theme.setDisabled(!canShape);
			stopsContainer.appendChild(theme.el);
			return { color, theme, position, remove };
		});
	};

	const gated = [fill, stroke, fillOpacity, gradientToggle, patternToggle];
	const gradientGated = [linearBtn, radialBtn, angleField, addStopBtn];
	const patternGated = [patternPreset, patternBackground];
	const effects = createShapeEffectsControls(doc, t, handlers);
	el.appendChild(effects.el);

	return {
		el,
		update(state) {
			el.hidden = !state.hasSelection || !state.canShape;
			fill.setValue(state.fillColor);
			stroke.setValue(state.strokeColor);
			fillOpacity.setValue(state.fillOpacity);
			lineFields.update(state.shapeStyle, !state.canShape);
			fillTheme.setThemeColorMap(state.themeColorMap);
			fillTheme.setSelected(state.fillColorRef, state.fillColor);
			fillTheme.setDisabled(!state.canShape);
			strokeTheme.setThemeColorMap(state.themeColorMap);
			strokeTheme.setSelected(state.strokeColorRef, state.strokeColor);
			strokeTheme.setDisabled(!state.canShape);
			fillRecent.setColors(state.recentColors ?? []);
			strokeRecent.setColors(state.recentColors ?? []);
			fillRecent.setDisabled(!state.canShape);
			strokeRecent.setDisabled(!state.canShape);
			gradientToggle.setValue(state.gradientEnabled);
			angleField.setValue(state.gradient.angle);
			lastGradient = state.gradient;
			rebuildStopRows(state.gradient.stops, state.themeColorMap, state.canShape);
			gradientPanel.hidden = !state.gradientEnabled;
			effects.update(state);

			const patternEnabled = state.shapeStyle?.fillMode === 'pattern';
			patternToggle.setValue(patternEnabled);
			lastPatternPreset = state.shapeStyle?.fillPatternPreset ?? lastPatternPreset;
			lastPatternBackground = state.shapeStyle?.fillPatternBackgroundColor ?? lastPatternBackground;
			patternPreset.setValue(lastPatternPreset);
			patternBackground.setValue(lastPatternBackground);
			patternPanel.hidden = !patternEnabled;

			for (const c of gated) {
				c.setDisabled(!state.canShape);
			}
			for (const c of gradientGated) {
				c.setDisabled(!state.canShape || !state.gradientEnabled);
			}
			for (const c of patternGated) {
				c.setDisabled(!state.canShape || !patternEnabled);
			}
			for (const row of stopRows) {
				row.color.setDisabled(!state.canShape);
				row.position.setDisabled(!state.canShape);
				row.remove.setDisabled(!state.canShape || state.gradient.stops.length <= 2);
			}
		},
	};
}
