import type { AutomationFormat, AutomationTool } from './types';
import { docxTools } from './docx-tools';
import { visioTools } from './visio-tools';
import { xlsxTools } from './xlsx-tools';

export type {
	AutomationCreateTool,
	AutomationEditTool,
	AutomationFormat,
	AutomationReadTool,
	AutomationTool,
} from './types';
export { UNTRUSTED_NOTE } from './types';
export { docxTools, visioTools, xlsxTools };

/** Every headless tool, by format. PowerPoint's tools are in `ooxml-core/pptx/automation`. */
export const automationTools: Readonly<Record<AutomationFormat, readonly AutomationTool[]>> = {
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
