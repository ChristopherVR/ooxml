import {
	resolveShapeFillView,
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillPatch,
} from 'ooxml-ui/pptx';
import type { ShapeFillAction, ShapeFillView } from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { makeButton, makeColorControl, makeNumberField } from '../controls';
import { createRecentColorsRow } from '../recent-colors-row';
import { createThemeColorSwatchGrid } from '../theme-color-swatch-grid';
import { createInspectorFields } from './inspector-fields';
import type { InspectorHandlers, InspectorState } from './types';

export interface ShapeFillControls {
	update(state: InspectorState): void;
}

/**
 * One labelled group per colour target, as PowerPoint's Format Shape pane
 * separates Fill from Line: the picker, then that target's theme grid and
 * recent colours.
 */
export function appendColorGroup(
	doc: Document,
	target: HTMLElement,
	label: string,
	control: HTMLElement,
	...rest: HTMLElement[]
): HTMLElement {
	const group = createEl(doc, 'div', 'pptxv-inspector-color-group');
	group.setAttribute('role', 'group');
	group.setAttribute('aria-label', label);
	const head = createEl(doc, 'div', 'pptxv-inspector-row');
	const title = createEl(doc, 'span', 'pptxv-inspector-row-label');
	title.textContent = label;
	head.append(title, control);
	group.append(head, ...rest);
	target.appendChild(group);
	return group;
}

/**
 * The Format pane's Fill section: the shared fill descriptors drawn by the
 * generic renderer (kind, transparency, gradient type and angle, picture
 * tiling), and the compound controls `resolveShapeFillView` asks for (solid
 * colour, gradient stops, pattern swatches and colours, picture source), whose
 * events become the shared `shapeFillPatch` actions through `setShapeStyle`.
 * The solid colour keeps the live colour actions (`editor-live-colors.ts`).
 */
export function createShapeFillControls(
	doc: Document,
	t: Translator,
	el: HTMLElement,
	handlers: InspectorHandlers,
): ShapeFillControls {
	let style: InspectorState['shapeStyle'];
	const act = (action: ShapeFillAction): void => {
		const patch = shapeFillPatch(style, action);
		if (patch) {
			handlers.setShapeStyle(patch);
		}
	};

	const modeFields = createInspectorFields(doc, t, {
		sections: SHAPE_FILL_MODE_SECTIONS,
		onPatch: (patch) => handlers.setShapeStyle(patch),
	});
	el.appendChild(modeFields.el);

	// The native picker reports every `input`, `change` and `blur` to the live colour actions.
	const fill = makeColorControl(
		doc,
		{
			label: t('pptx.inspector.fill'),
			onInput: (hex) => handlers.liveColor('fill', 'input', hex),
			onCommit: (hex) => handlers.liveColor('fill', 'change', hex),
			onBlur: (hex) => handlers.liveColor('fill', 'blur', hex),
		},
		'#4f86ff',
	);
	// A theme swatch commits BOTH the resolved hex and the ref; the recent row clears it.
	const fillTheme = createThemeColorSwatchGrid(doc, t, (commit) =>
		handlers.setShapeFill(commit.hex, commit.ref),
	);
	const fillRecent = createRecentColorsRow(doc, t, handlers.setShapeFill);
	const colorGroup = appendColorGroup(
		doc,
		el,
		t('pptx.inspector.fill'),
		fill.el,
		fillTheme.el,
		fillRecent.el,
	);

	const optionFields = createInspectorFields(doc, t, {
		sections: SHAPE_FILL_OPTIONS_SECTIONS,
		onPatch: (patch) => handlers.setShapeStyle(patch),
	});
	el.appendChild(optionFields.el);

	const gradientPanel = createEl(doc, 'div', 'pptxv-inspector-gradient');
	const stopsContainer = createEl(doc, 'div', 'pptxv-inspector-gradient-stops');
	const addStop = makeButton(doc, {
		label: t('pptx.gradient.addStop'),
		text: t('pptx.gradient.addStop'),
		onClick: () => act({ type: 'addStop' }),
	});
	gradientPanel.append(stopsContainer, addStop.btn);
	el.appendChild(gradientPanel);

	const patternPanel = createEl(doc, 'div', 'pptxv-inspector-pattern');
	const patternGrid = createEl(doc, 'div', 'pptxv-inspector-pattern-grid');
	patternGrid.setAttribute('role', 'group');
	patternGrid.setAttribute('aria-label', t('pptx.table.patternPreset'));
	const patternColor = (type: 'patternForeground' | 'patternBackground', label: string) =>
		makeColorControl(doc, {
			label,
			onInput: (hex) => act({ type, hex }),
			onCommit: handlers.pushRecentColor,
		});
	const foreground = patternColor('patternForeground', t('pptx.fillAdvanced.foregroundColor'));
	const background = patternColor('patternBackground', t('pptx.fillAdvanced.backgroundColor'));
	patternPanel.append(patternGrid, foreground.el, background.el);
	el.appendChild(patternPanel);

	const imagePanel = createEl(doc, 'div', 'pptxv-inspector-pattern');
	const imageUrl = doc.createElement('input');
	imageUrl.type = 'text';
	imageUrl.className = 'pptxv-field-input';
	imageUrl.dataset.testid = 'fill-image-url';
	imageUrl.placeholder = 'https://example.com/image.png';
	imageUrl.setAttribute('aria-label', t('pptx.fillAdvanced.imageUrl'));
	imageUrl.addEventListener('change', () => act({ type: 'imageUrl', url: imageUrl.value }));
	const imageFile = doc.createElement('input');
	imageFile.type = 'file';
	imageFile.accept = 'image/*';
	imageFile.setAttribute('aria-label', t('pptx.fillAdvanced.orChooseFile'));
	imageFile.addEventListener('change', () => {
		const file = imageFile.files?.[0];
		if (!file) {
			return;
		}
		const reader = new FileReader();
		reader.onload = () => act({ type: 'imageUrl', url: String(reader.result) });
		reader.readAsDataURL(file);
	});
	imagePanel.append(imageUrl, imageFile);
	el.appendChild(imagePanel);

	const drawStops = (
		stops: NonNullable<ShapeFillView['stops']>,
		state: InspectorState,
		disabled: boolean,
	): void => {
		stopsContainer.replaceChildren();
		for (const stop of stops) {
			const row = createEl(doc, 'div', 'pptxv-inspector-row');
			row.dataset.testid = 'fill-gradient-stop';
			const color = makeColorControl(
				doc,
				{
					label: t('pptx.gradient.stops'),
					onInput: (hex) => act({ type: 'stopColor', index: stop.index, hex }),
					onCommit: handlers.pushRecentColor,
				},
				stop.color,
			);
			const position = makeNumberField(doc, {
				label: t('pptx.gradient.position'),
				min: 0,
				max: 100,
				onCommit: (value) => act({ type: 'stopPosition', index: stop.index, position: value }),
			});
			position.setValue(stop.position);
			const opacity = makeNumberField(doc, {
				label: t('pptx.strokeEffects.fillOpacity'),
				min: 0,
				max: 100,
				onCommit: (value) => act({ type: 'stopOpacity', index: stop.index, opacity: value }),
			});
			opacity.setValue(stop.opacity);
			const remove = makeButton(doc, {
				label: t('pptx.gradient.removeStop'),
				text: '✕',
				onClick: () => act({ type: 'removeStop', index: stop.index }),
			});
			color.setDisabled(disabled);
			position.setDisabled(disabled);
			opacity.setDisabled(disabled);
			remove.setDisabled(disabled || !stop.removable);
			row.append(color.el, position.el, opacity.el, remove.btn);
			const theme = createThemeColorSwatchGrid(doc, t, (commit) =>
				act({ type: 'stopColor', index: stop.index, hex: commit.hex, ref: commit.ref }),
			);
			theme.setThemeColorMap(state.themeColorMap);
			theme.setSelected(stop.colorRef, stop.color);
			theme.setDisabled(disabled);
			stopsContainer.append(row, theme.el);
		}
	};

	const drawPattern = (pattern: NonNullable<ShapeFillView['pattern']>, disabled: boolean) => {
		patternGrid.replaceChildren();
		for (const swatch of pattern.swatches) {
			const button = createEl(doc, 'button', 'pptxv-inspector-pattern-swatch');
			button.type = 'button';
			button.dataset.testid = 'fill-pattern-swatch';
			button.title = t(swatch.labelKey);
			button.setAttribute('aria-label', t(swatch.labelKey));
			button.setAttribute('aria-pressed', String(swatch.selected));
			button.disabled = disabled;
			if (swatch.backgroundImage) {
				button.style.backgroundImage = swatch.backgroundImage;
			}
			button.addEventListener('click', () => act({ type: 'patternPreset', preset: swatch.value }));
			patternGrid.appendChild(button);
		}
		foreground.setValue(pattern.foreground);
		background.setValue(pattern.background);
		foreground.setDisabled(disabled);
		background.setDisabled(disabled);
	};

	return {
		update(state) {
			style = state.shapeStyle;
			const disabled = !state.canShape;
			const view = resolveShapeFillView(style);
			modeFields.update(style, disabled);
			optionFields.update(style, disabled);

			colorGroup.hidden = !view.color;
			fill.setValue(state.fillColor);
			fill.setDisabled(disabled);
			fillTheme.setThemeColorMap(state.themeColorMap);
			fillTheme.setSelected(state.fillColorRef, state.fillColor);
			fillTheme.setDisabled(disabled);
			fillRecent.setColors(state.recentColors ?? []);
			fillRecent.setDisabled(disabled);

			gradientPanel.hidden = !view.stops;
			addStop.setDisabled(disabled);
			drawStops(view.stops ?? [], state, disabled);

			patternPanel.hidden = !view.pattern;
			if (view.pattern) {
				drawPattern(view.pattern, disabled);
			} else {
				patternGrid.replaceChildren();
			}

			imagePanel.hidden = !view.image;
			imageUrl.value = view.image?.url ?? '';
			imageUrl.disabled = disabled;
			imageFile.disabled = disabled;
		},
	};
}
