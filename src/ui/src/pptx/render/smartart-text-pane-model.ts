/**
 * The PowerPoint SmartArt text pane as a pure model: the outline lines, notes and append state
 * `pptx-ui-smartart-text-pane` shows, and the edit each of its requests makes to the data. Every
 * edit goes through the core node operations (the pptx `smartart-editing` facades over
 * `diagram/edit`) and the shared pane handlers, so every binding edits a diagram the same way.
 *
 * @module render/smartart-text-pane-model
 */
import { diagramNodeLevels } from 'ooxml-core/diagram';
import type { PptxSmartArtData, PptxSmartArtNodeStyle } from 'ooxml-core/pptx';
import {
	addSmartArtNodeAsChild,
	setSmartArtNodeStyle,
	updateSmartArtNodeText,
} from 'ooxml-core/pptx';

import {
	canAddTopLevelNode,
	canRemoveTopLevelNode,
	describeSmartArtBounds,
} from './smartart-node-limits';
import {
	addSiblingAfter,
	classifyExtraConnections,
	demote,
	promote,
	removeEmptyNode,
	reorder,
} from './smartart-node-pane-handlers';

type Translate = (key: string, params?: Record<string, string | number>) => string;

/** One outline line, in the shape `office-ui-smartart-text-pane` takes. */
export interface SmartArtTextPaneLine {
	id: string;
	text: string;
	level: number;
	/** The item number of a top-level line, a bullet for a deeper one. */
	marker: string;
	label: string;
	style?: PptxSmartArtNodeStyle;
	canMoveUp: boolean;
	canMoveDown: boolean;
	canRemove: boolean;
	/** Sub-items are added under top-level lines only. */
	canAddChild: boolean;
}

/** What the pane shows for a diagram. */
export interface SmartArtTextPaneView {
	lines: SmartArtTextPaneLine[];
	/** The layout holds no more top-level shapes. */
	appendDisabled: boolean;
	/** The layout's shape-count hint, the extra connections note and the keyboard hint. */
	notes: string[];
}

/** A request from the pane, as the pptx alias reads it from the element's events. */
export type SmartArtTextPaneRequest =
	| { kind: 'text'; id: string; text: string }
	| { kind: 'add' | 'remove' | 'promote' | 'demote' | 'addChild'; id: string }
	| { kind: 'move'; id: string; direction: 1 | -1 }
	| { kind: 'append' }
	| { kind: 'style'; id: string; style: Partial<PptxSmartArtNodeStyle> };

/** The edited data and the line to focus after it lands. */
export interface SmartArtTextPaneEdit {
	data: PptxSmartArtData;
	focusNodeId?: string;
}

const BULLET = '•';

function topLevelCount(levels: readonly number[]): number {
	return levels.filter((level) => level === 0).length;
}

/** The outline, notes and append state for `data`. */
export function smartArtTextPaneView(data: PptxSmartArtData, t: Translate): SmartArtTextPaneView {
	const nodes = data.nodes ?? [];
	const levels = diagramNodeLevels(nodes);
	const tops = topLevelCount(levels);
	const layout = data.resolvedLayoutType;
	const canRemoveTop = canRemoveTopLevelNode(layout, tops);
	let number = 0;
	const lines = nodes.map((node, index): SmartArtTextPaneLine => {
		const level = levels[index] ?? 0;
		const siblings = nodes.filter((other) => other.parentId === node.parentId);
		const position = siblings.indexOf(node);
		const text = node.text || t('pptx.smartArt.emptyNode');
		if (level === 0) {
			number += 1;
		}
		return {
			id: node.id,
			text: node.text,
			level,
			marker: level === 0 ? String(number) : BULLET,
			label:
				level === 0
					? t('pptx.smartArt.itemLabel', { index: number, text })
					: t('pptx.smartArt.subItemLabel', { text }),
			...(node.style ? { style: node.style } : {}),
			canMoveUp: position > 0,
			canMoveDown: position >= 0 && position < siblings.length - 1,
			canRemove: nodes.length > 1 && (level > 0 || canRemoveTop),
			canAddChild: level === 0,
		};
	});
	const notes: string[] = [];
	const bounds = describeSmartArtBounds(layout);
	if (bounds) {
		notes.push(bounds);
	}
	const extra = classifyExtraConnections(data);
	if (extra.length > 0) {
		const types = [...new Set(extra.map((c) => c.type ?? 'unknown'))].join(', ');
		notes.push(`${t('pptx.smartart.extraConnections', { count: extra.length })} (${types})`);
	}
	notes.push(t('pptx.smartart.tabHint'));
	return { lines, appendDisabled: !canAddTopLevelNode(layout, tops), notes };
}

/** The id of the node `next` has and `data` does not (the one an add created). */
function addedNodeId(data: PptxSmartArtData, next: PptxSmartArtData): string | undefined {
	const ids = new Set(data.nodes.map((node) => node.id));
	return next.nodes.find((node) => !ids.has(node.id))?.id;
}

function canRemove(data: PptxSmartArtData, id: string): boolean {
	const nodes = data.nodes;
	const index = nodes.findIndex((node) => node.id === id);
	if (index < 0 || nodes.length <= 1) {
		return false;
	}
	const levels = diagramNodeLevels(nodes);
	return (
		levels[index] !== 0 || canRemoveTopLevelNode(data.resolvedLayoutType, topLevelCount(levels))
	);
}

function withFocus(
	next: PptxSmartArtData | undefined,
	id: string,
): SmartArtTextPaneEdit | undefined {
	return next ? { data: next, focusNodeId: id } : undefined;
}

/**
 * The edit a pane request makes, or `undefined` when it changes nothing or the layout refuses it
 * (removing a shape the layout needs, appending past its maximum). `subItemText` labels a new
 * sub-item.
 */
export function applySmartArtTextPaneRequest(
	data: PptxSmartArtData,
	request: SmartArtTextPaneRequest,
	subItemText: string,
): SmartArtTextPaneEdit | undefined {
	switch (request.kind) {
		case 'text': {
			const node = data.nodes.find((n) => n.id === request.id);
			return node && node.text !== request.text
				? { data: updateSmartArtNodeText(data, request.id, request.text) }
				: undefined;
		}
		case 'add': {
			const result = addSiblingAfter(data, request.id);
			return result
				? { data: result.data, ...(result.focusNodeId ? { focusNodeId: result.focusNodeId } : {}) }
				: undefined;
		}
		case 'remove': {
			if (!canRemove(data, request.id)) {
				return undefined;
			}
			const result = removeEmptyNode(data, request.id);
			return result
				? { data: result.data, ...(result.focusNodeId ? { focusNodeId: result.focusNodeId } : {}) }
				: undefined;
		}
		case 'promote':
			return withFocus(promote(data, request.id), request.id);
		case 'demote':
			return withFocus(demote(data, request.id), request.id);
		case 'move':
			return withFocus(reorder(data, request.id, request.direction), request.id);
		case 'addChild': {
			const next = addSmartArtNodeAsChild(data, request.id, subItemText);
			const id = addedNodeId(data, next);
			return id ? { data: next, focusNodeId: id } : undefined;
		}
		case 'append': {
			const levels = diagramNodeLevels(data.nodes);
			if (!canAddTopLevelNode(data.resolvedLayoutType, topLevelCount(levels))) {
				return undefined;
			}
			const next = addSmartArtNodeAsChild(data);
			const id = addedNodeId(data, next);
			return id ? { data: next, focusNodeId: id } : undefined;
		}
		case 'style': {
			const next = setSmartArtNodeStyle(data, request.id, request.style);
			return next === data ? undefined : { data: next };
		}
	}
}
