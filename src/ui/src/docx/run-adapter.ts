import type { TextRun } from 'ooxml-core/docx';
import {
	runToInlineNodes as coreRunToInlineNodes,
	type NoteNumberLookup,
} from 'ooxml-core/docx/ui';
import { schema } from './schema';
export { appendInlineNode } from 'ooxml-core/docx/ui';
export type { NoteNumberLookup } from 'ooxml-core/docx/ui';

export const runToInlineNodes = (run: TextRun, noteNumber?: NoteNumberLookup) =>
	coreRunToInlineNodes(run, schema, noteNumber);
