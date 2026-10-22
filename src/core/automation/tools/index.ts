import type { AutomationFormat, AutomationTool } from './types';
import { docxTools } from './docx-tools';
import { visioTools } from './visio-tools';
import { xlsxTools } from './xlsx-tools';

export type {
	AutomationCreateTool,
	AutomationEditResult,
	AutomationEditTool,
	AutomationFormat,
	AutomationReadTool,
	AutomationTool,
} from './types';
export { UNTRUSTED_NOTE } from './types';
export { docxTools, visioTools, xlsxTools };

/**
 * The Word, Excel and Visio tools, by format. PowerPoint's are `pptxTools` in
 * `ooxml-core/pptx/automation` (the same kind of definitions, kept out of this entry so it never
 * pulls in the PowerPoint bundle).
 */
export const automationTools: Readonly<
	Record<Exclude<AutomationFormat, 'pptx'>, readonly AutomationTool[]>
> = {
	docx: docxTools,
	xlsx: xlsxTools,
	visio: visioTools,
};

/** The tool called `name`, in any format. */
export function findAutomationTool(name: string): AutomationTool | undefined {
	for (const tools of Object.values(automationTools)) {
		const tool = tools.find((candidate) => candidate.name === name);
		if (tool) return tool;
	}
	return undefined;
}
