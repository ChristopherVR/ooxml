import { z } from 'zod';
import { editVsdx } from '../../visio/edit';
import type { VisioEdit } from '../../visio/index';
import { inspectVisio } from '../visio';
import { editTool, id, longText, readTool } from './define';
import type { AutomationTool } from './types';

const format = 'visio';
const extensions = ['.vsdx'];
const target = { pageId: id, shapeId: id };
const finite = z.number().finite();

const edit = z.discriminatedUnion('type', [
	z.strictObject({ ...target, type: z.literal('replace-plain-text'), text: longText }),
	z.strictObject({ ...target, type: z.literal('move-shape'), x: finite, y: finite }),
	z.strictObject({
		...target,
		type: z.literal('resize-shape'),
		width: finite.positive(),
		height: finite.positive(),
	}),
	z.strictObject({ ...target, type: z.literal('delete-shape') }),
	z.strictObject({
		...target,
		type: z.literal('create-rectangle'),
		x: finite,
		y: finite,
		width: finite.positive(),
		height: finite.positive(),
		text: longText.optional(),
	}),
]);

/** Visio tools: inspect pages and shapes, and apply a batch of the supported edits. */
export const visioTools: readonly AutomationTool[] = [
	readTool({
		name: 'visio_inspect',
		description: 'Inspect VSDX pages, shape ids, text, geometry and core diagnostics.',
		format,
		extensions,
		input: {},
		run: (bytes) => inspectVisio(bytes),
	}),
	editTool({
		name: 'visio_edit',
		description:
			'Apply an atomic batch of core-supported experimental Visio text and geometry edits. Coordinates use inches with a bottom-left origin. Unsupported edits reject; diagnostics report fidelity limits.',
		format,
		extensions,
		input: { edits: z.array(edit).min(1).max(1000) },
		// The schema admits a subset of VisioEdit; zod leaves absent optional keys absent.
		run: (bytes, p) => editVsdx(bytes, p.edits as readonly VisioEdit[]),
	}),
];
