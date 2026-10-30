import type { VisioEdit } from '../edit-commands';
import type { VisioPage } from '../model';
import { visioShapePageBounds } from './diagram-parts';
import { visioNextShapeId } from './shape-id';
import { visioSwimlaneInfo } from './swimlane';
import {
	visioVisualizerDepths,
	visioVisualizerSize,
	type VisioVisualizerStep,
} from './data-visualizer';

/** Hidden Shape Data row on every shape the visualizer made: label step or link, value its key. */
export const VISIO_VISUALIZER_ROW = '_ooxmlVisualizer';
export type VisioVisualizerVariant = 'basic' | 'cross-functional';

const LINE = '#326765';
const FILLS: Record<string, string> = {
	'flowchart-terminator': '#F3EEE1',
	'flowchart-decision': '#DFEBF7',
};
const COLUMN = 2.6;
const LANE = 1.4;
const TITLE = 0.35;
const LEFT = 0.6;
const GUTTER = 1.5;

const rowName = (name: string) => name.replace(/[^A-Za-z0-9]/g, '') || 'Column';

/** The shapes this tool made on a page, by step ID, plus its connectors. */
export function visioVisualizerShapes(page: VisioPage) {
	const steps = new Map<string, string>();
	const links: string[] = [];
	for (const shape of page.shapes) {
		const row = shape.shapeData?.find((item) => item.name === VISIO_VISUALIZER_ROW);
		if (!row) continue;
		if (row.label === 'step') steps.set(String(row.value ?? row.rawValue ?? ''), shape.id);
		else links.push(shape.id);
	}
	return { steps, links };
}

const hidden = (label: string, value: string, invisible = true) => ({
	label,
	prompt: '',
	type: 'string' as const,
	format: '',
	value,
	invisible,
});

/**
 * Edits that draw (or redraw) a flowchart from steps. Shapes this tool made earlier are removed
 * and drawn again from the table, each step that was already there keeping its place; shapes
 * the user added are left alone. A cross-functional flowchart adds its swimlanes when the page
 * has none and otherwise places steps in the lane whose heading equals the step's Function.
 */
export function visioDataVisualizerEdits(
	page: VisioPage,
	steps: readonly VisioVisualizerStep[],
	variant: VisioVisualizerVariant,
): { batches: VisioEdit[][]; notes: string[] } {
	const ratio = page.drawingToPageScale ?? 1;
	const old = visioVisualizerShapes(page);
	const centre = new Map<string, { x: number; y: number }>();
	for (const [stepId, shapeId] of old.steps) {
		const shape = page.shapes.find((item) => item.id === shapeId);
		if (shape?.rotation)
			centre.set(stepId, { x: shape.rotation.pinX / ratio, y: shape.rotation.pinY / ratio });
	}
	const removals: VisioEdit[] = [...old.links, ...old.steps.values()].map((shapeId): VisioEdit => ({
		type: 'delete-shape',
		pageId: page.id,
		shapeId,
	}));
	const edits: VisioEdit[] = [];
	const notes: string[] = [];
	const pageEdits: VisioEdit[] = [];
	const depth = visioVisualizerDepths(steps);
	const columns = Math.max(...depth.values()) + 1;
	const frame =
		variant === 'cross-functional'
			? page.shapes.find(
					(shape) =>
						shape.structure?.type === 'container' && shape.structure.swimlane?.role === 'frame',
				)
			: undefined;
	const info = frame ? visioSwimlaneInfo(page, frame.id) : undefined;
	const names = [...new Set(steps.map((step) => step.lane || 'Unassigned'))];
	let next = Number(visioNextShapeId(page));
	const id = () => String(next++);
	const grow = (width: number, height: number) => {
		if (width * ratio > page.width || height * ratio > page.height)
			pageEdits.push({
				type: 'set-page-size',
				pageId: page.id,
				width: Math.max(width * ratio, page.width),
				height: Math.max(height * ratio, page.height),
			});
	};
	// Where each step goes (drawing inches, y up).
	const at = new Map<string, { x: number; y: number }>();
	if (variant === 'basic') {
		const layers = new Map<number, VisioVisualizerStep[]>();
		for (const step of steps) {
			const layer = depth.get(step.id)!;
			layers.set(layer, [...(layers.get(layer) ?? []), step]);
		}
		const widest = Math.max(...[...layers.values()].map((members) => members.length));
		const width = Math.max(page.width / ratio, widest * 2.8 + 1);
		const height = Math.max(page.height / ratio, layers.size * 1.5 + 1);
		grow(width, height);
		for (const [layer, members] of layers)
			members.forEach((step, index) =>
				at.set(
					step.id,
					centre.get(step.id) ?? {
						x: width / 2 + (index - (members.length - 1) / 2) * 2.8,
						y: height - 0.9 - layer * 1.5,
					},
				),
			);
		// A new step goes under a step that kept its place, so it never lands on another one.
		for (const step of steps) {
			if (centre.has(step.id)) continue;
			const parents = steps.filter(
				(item) => centre.has(item.id) && item.next.some((link) => link.id === step.id),
			);
			if (!parents.length) continue;
			const parent = centre.get(parents[0]!.id)!;
			const taken = [...at.values()];
			let x = parent.x;
			while (
				taken.some((spot) => Math.abs(spot.x - x) < 2.4 && Math.abs(spot.y - (parent.y - 1.5)) < 1)
			)
				x += 2.8;
			at.set(step.id, { x, y: parent.y - 1.5 });
		}
	} else if (info) {
		const left = visioShapePageBounds(info.frame).minX / ratio;
		const lanes = info.lanes.map((lane) => {
			const box = visioShapePageBounds(lane);
			return { name: lane.text.plainText.trim(), y: (box.minY + box.maxY) / 2 / ratio };
		});
		for (const step of steps) {
			const wanted = step.lane || 'Unassigned';
			const lane = lanes.find((item) => item.name === wanted) ?? lanes[0]!;
			if (lane.name !== wanted)
				notes.push(`Step ${step.id}: there is no "${wanted}" lane; placed in "${lane.name}".`);
			at.set(step.id, {
				x: centre.get(step.id)?.x ?? left + GUTTER + depth.get(step.id)! * COLUMN - 0.4,
				y: lane.y,
			});
		}
	} else {
		const bottom = 0.6;
		const frameWidth = GUTTER + columns * COLUMN;
		const frameHeight = names.length * LANE + TITLE;
		grow(LEFT + frameWidth + 0.6, bottom + frameHeight + 0.6);
		const frameId = id();
		const laneIds = names.map(id);
		edits.push({
			type: 'create-swimlanes',
			pageId: page.id,
			frameId,
			laneIds,
			headings: names,
			listId: id(),
			phaseListId: id(),
			box: {
				x: LEFT + frameWidth / 2,
				y: bottom + frameHeight / 2,
				width: frameWidth,
				height: frameHeight,
			},
			direction: 'horizontal',
			title: 'Cross-Functional Flowchart',
		});
		for (const step of steps)
			at.set(step.id, {
				x: LEFT + GUTTER + depth.get(step.id)! * COLUMN - 0.4,
				y: bottom + frameHeight - TITLE - (names.indexOf(step.lane || 'Unassigned') + 0.5) * LANE,
			});
	}
	const shapeIds = new Map(steps.map((step) => [step.id, id()]));
	for (const step of steps) {
		const shapeId = shapeIds.get(step.id)!;
		const place = at.get(step.id)!;
		edits.push(
			{
				type: 'drop-stencil-master',
				pageId: page.id,
				shapeId,
				master: step.master,
				x: place.x,
				y: place.y,
			},
			{ type: 'resize-shape', pageId: page.id, shapeId, ...visioVisualizerSize(step.master) },
			{ type: 'replace-plain-text', pageId: page.id, shapeId, text: step.description },
			{
				type: 'format-shape',
				pageId: page.id,
				shapeId,
				fillColor: FILLS[step.master] ?? '#D8EFEB',
				lineColor: LINE,
			},
			...step.values.map((value): VisioEdit => ({
				type: 'set-shape-data',
				pageId: page.id,
				shapeId,
				row: rowName(value.name),
				data: hidden(value.label, value.value, false),
			})),
			{
				type: 'set-shape-data',
				pageId: page.id,
				shapeId,
				row: VISIO_VISUALIZER_ROW,
				data: hidden('step', step.id),
			},
		);
	}
	for (const step of steps)
		for (const link of step.next) {
			const from = at.get(step.id)!;
			const to = at.get(link.id)!;
			const shapeId = id();
			edits.push(
				{
					type: 'create-line',
					pageId: page.id,
					shapeId,
					beginX: from.x,
					beginY: from.y,
					endX: to.x,
					endY: to.y,
					connect: { begin: shapeIds.get(step.id)!, end: shapeIds.get(link.id)! },
				},
				{ type: 'format-shape', pageId: page.id, shapeId, lineColor: LINE },
				{
					type: 'set-shape-data',
					pageId: page.id,
					shapeId,
					row: VISIO_VISUALIZER_ROW,
					data: hidden('link', `${step.id}>${link.id}`),
				},
			);
			if (link.label) {
				// A connector has no text block of its own here, so its label is a text box that
				// sits at the middle of the line (it does not follow the connector when a step moves).
				const labelId = id();
				edits.push(
					{
						type: 'create-text-box',
						pageId: page.id,
						shapeId: labelId,
						x: (from.x + to.x) / 2 + 0.05,
						y: (from.y + to.y) / 2,
						width: 0.7,
						height: 0.25,
						text: link.label,
					},
					{
						type: 'set-shape-data',
						pageId: page.id,
						shapeId: labelId,
						row: VISIO_VISUALIZER_ROW,
						data: hidden('link', `${step.id}>${link.id}:label`),
					},
				);
			}
		}
	if (variant === 'cross-functional')
		edits.push({ type: 'reconcile-containers', pageId: page.id, shapeIds: [...shapeIds.values()] });
	// Page, delete and draw edits cannot share a transaction, so they are applied in turn.
	return { batches: [pageEdits, removals, edits].filter((batch) => batch.length), notes };
}
