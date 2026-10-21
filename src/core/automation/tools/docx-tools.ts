import { z } from 'zod';
import { createDocx, inspectDocx, setDocxRunText, suggestDocxRunText } from '../docx';
import { fillDocxFormFields, listDocxFormFields } from '../docx-forms';
import { createTool, editTool, id, index, longText, readTool } from './define';
import type { AutomationTool } from './types';

const format = 'docx';
const extensions = ['.docx'];

/** Word tools: inspect, create, edit runs directly or as tracked changes, and fill forms. */
export const docxTools: readonly AutomationTool[] = [
	readTool({
		name: 'docx_inspect',
		description: 'Inspect Word blocks, stable paragraph ids, runs, properties and warnings.',
		format,
		extensions,
		input: {},
		run: (bytes) => inspectDocx(bytes),
	}),
	readTool({
		name: 'docx_list_form_fields',
		description:
			'List the form fields (content controls and legacy fields) with their ids, kinds, current values and choices, and any editing restriction.',
		format,
		extensions,
		input: {},
		run: (bytes) => listDocxFormFields(bytes),
	}),
	createTool({
		name: 'docx_create',
		description: 'Create a Word document with one paragraph per string.',
		format,
		extensions,
		input: { paragraphs: z.array(longText).max(10_000).default([]) },
		run: (p) => createDocx(p.paragraphs),
	}),
	editTool({
		name: 'docx_set_run_text',
		description:
			'Set one ordinary text run, preserving its formatting. Fields, images, equations and revisions are rejected.',
		format,
		extensions,
		input: { paragraphId: id, runIndex: index, text: longText },
		run: (bytes, p) => setDocxRunText(bytes, p.paragraphId, p.runIndex, p.text),
	}),
	editTool({
		name: 'docx_suggest_run_text',
		description:
			'Propose new text for ordinary runs as tracked changes (a deletion and an insertion by author) that the user accepts or rejects in Word. Run indexes refer to the document before these suggestions.',
		format,
		extensions,
		input: {
			suggestions: z
				.array(z.strictObject({ paragraphId: id, runIndex: index, text: longText }))
				.min(1)
				.max(1000),
			author: z.string().min(1).max(255).optional(),
		},
		run: (bytes, p) =>
			suggestDocxRunText(bytes, p.suggestions, p.author ? { author: p.author } : {}),
	}),
	editTool({
		name: 'docx_fill_form_fields',
		description:
			'Fill form fields by the ids docx_list_form_fields gives: text, a checkbox boolean, a choice, or a YYYY-MM-DD date. If any value is refused nothing is written, unless partial is true.',
		format,
		extensions,
		input: {
			values: z.record(id, z.union([longText, z.boolean(), z.number().finite()])),
			partial: z.boolean().optional(),
		},
		run: (bytes, p) => fillDocxFormFields(bytes, p.values, p.partial ? { partial: true } : {}),
	}),
];
