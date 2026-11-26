import { findDiagramNativeColor } from 'ooxml-core/diagram';
import type { PptxSmartArtData } from 'ooxml-core/pptx';

import { smartArtColorPatch } from './ribbon-galleries/smartart-color-patch';
import { SMARTART_COLOR_ENTRIES } from './ribbon-galleries/smartart-gallery-patch';
import {
	appliedSmartArtColor,
	modelColorScheme,
	smartArtColorLabel,
} from './ribbon-galleries/smartart-native-entries';
import { schemaLabel, SMARTART_COLOR_SCHEME_LABEL_KEYS } from './schema-label-keys';
import { PALETTES } from './smartart-drawing';

type Translate = (key: string, params?: Record<string, string | number>) => string;

/** The inspector edits model schemes and also displays the active ribbon-only definition. */
export function smartArtInspectorColors(data: PptxSmartArtData | undefined, t: Translate) {
	const entry = data && findDiagramNativeColor(appliedSmartArtColor(data));
	const options: { value: string; label: string }[] = SMARTART_COLOR_ENTRIES.map(({ id }) => ({
		value: id,
		label: schemaLabel(SMARTART_COLOR_SCHEME_LABEL_KEYS, id, t),
	}));
	const model = entry && modelColorScheme(entry);
	if (entry && !model) {
		const { labelKey, labelParams } = smartArtColorLabel(entry);
		options.push({ value: entry.id, label: t(labelKey, labelParams) });
		return { value: entry.id, options };
	}
	const custom = data?.colorTransform?.name;
	if (!entry && custom) {
		options.push({ value: 'custom', label: custom });
		return { value: 'custom', options };
	}
	return { value: model ?? data?.colorScheme ?? 'colorful1', options };
}

/** A model scheme selected in the inspector replaces the previous definition and its fills. */
export function setSmartArtInspectorColor(data: PptxSmartArtData, value: string): PptxSmartArtData {
	const scheme = SMARTART_COLOR_ENTRIES.find(({ id }) => id === value)?.id;
	if (!scheme || scheme === smartArtInspectorColors(data, (key) => key).value) return data;
	const patch =
		data.colorTransform || data.drawingShapes?.length
			? smartArtColorPatch(data, PALETTES[scheme], `urn:ooxml-core/colors/${scheme}`)
			: { drawingShapes: undefined };
	return { ...data, ...patch, colorScheme: scheme, builtinColorsXml: undefined };
}
