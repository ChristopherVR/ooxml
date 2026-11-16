/**
 * Parse and serialize `a:tabLst/a:tab` tab stops.
 *
 * One implementation for every reader (slide paragraphs, placeholder level
 * styles, SmartArt text) and every writer (the paragraph writer and the
 * placeholder level-style serializer), so the default-value rules cannot
 * drift between them. The attributes are read and written by the shared
 * DrawingML codec (`readDrawingTabStop`, `drawingTabStopEntries`); this
 * module converts EMU to px and keeps which schema defaults were authored.
 *
 * @module tab-stops
 */

import type { DrawingTabStop } from '../../../drawingml/text-model';
import { drawingTabStopEntries, readDrawingTabStop } from '../../../drawingml/text-paragraph-codec';
import type { TextStyle, XmlObject } from '../types';
import { objectAttributes, objectTextElementView } from './paragraph-properties-neutral';

const EMU_PER_PX = 9525;

type TabStop = NonNullable<TextStyle['tabStops']>[number];

const ALIGNS = new Set(['ctr', 'r', 'dec']);
const LEADERS = new Set(['dot', 'hyphen', 'underscore']);

/**
 * Parse tab stops from `a:pPr > a:tabLst > a:tab`.
 *
 * Each tab has `@_pos` (EMU), `@_algn`, and optional `@_leader`. Returns
 * `undefined` when there is no `a:tabLst` or it holds no `a:tab`.
 */
export function parseTabStops(pPr: XmlObject | undefined): TextStyle['tabStops'] | undefined {
	const tabs = objectTextElementView(pPr)?.child('tabLst')?.children('tab') ?? [];
	if (tabs.length === 0) {
		return undefined;
	}
	return tabs
		.filter((tab) => tab.attribute('pos') !== undefined)
		.map((tab) => tabStopFromDrawing(readDrawingTabStop(tab)));
}

/** A neutral tab stop as the pptx model keeps it (px, default tokens resolved). */
export function tabStopFromDrawing(stop: DrawingTabStop): TabStop {
	const position = (stop.positionEmu ?? 0) / EMU_PER_PX;
	const align = stop.align && ALIGNS.has(stop.align) ? (stop.align as TabStop['align']) : 'l';
	const leader =
		stop.leader && LEADERS.has(stop.leader) ? (stop.leader as TabStop['leader']) : undefined;
	// The schema defaults (`algn="l"`, `leader="none"`) are omitted by the
	// writer unless the source spelled them out; remember when it did.
	const alignAuthored = stop.align === 'l';
	const leaderAuthored = stop.leader === 'none';
	return {
		position,
		align,
		...(leader ? { leader } : {}),
		...(alignAuthored ? { alignAuthored } : {}),
		...(leaderAuthored ? { leaderAuthored } : {}),
	};
}

/** A pptx tab stop on the neutral model (px -> EMU, defaults only when authored). */
export function tabStopToDrawing(tab: TabStop): DrawingTabStop {
	const stop: DrawingTabStop = { positionEmu: Math.round(tab.position * EMU_PER_PX) };
	if (tab.align && (tab.align !== 'l' || tab.alignAuthored)) {
		stop.align = tab.align;
	}
	if (tab.leader && tab.leader !== 'none') {
		stop.leader = tab.leader;
	} else if (!tab.leader && tab.leaderAuthored) {
		stop.leader = 'none';
	}
	return stop;
}

/**
 * Serialize one typed tab stop to an `a:tab` node (positions px -> EMU).
 * Schema defaults (`algn="l"`, `leader="none"`) are written only when the
 * source authored them, so an unedited `a:tabLst` round-trips exactly.
 */
export function serializeTabStop(tab: TabStop): XmlObject {
	return objectAttributes(drawingTabStopEntries(tabStopToDrawing(tab)));
}
