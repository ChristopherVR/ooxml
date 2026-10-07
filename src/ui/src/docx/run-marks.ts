import type { TextRun } from 'ooxml-core/docx';
import { marksForRun as coreMarksForRun } from 'ooxml-core/docx/ui';
import { schema } from './schema';

export const marksForRun = (run: TextRun) => coreMarksForRun(run, schema);
