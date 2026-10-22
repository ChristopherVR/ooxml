import { editTool, readTool } from '../../../automation/tools/define';
import type { AutomationTool } from '../../../automation/tools/types';
import { arrayBufferOf } from '../execution';
import * as schemas from '../schemas/index';
import * as conversionTools from '../tools/conversion-tools';
import * as validationTools from '../tools/validation-tools';
import { DECK_EDIT_EXTENSIONS, DECK_READ_EXTENSIONS, deckTool } from './deck-tool';

/**
 * A relative path without `..` segments. The converter writes nothing here (it has no file
 * system), so `outputDir` only prefixes the media links in the Markdown; refusing absolute paths
 * keeps a host path out of the output.
 */
function relativeDir(dir: string | undefined): string {
	const value = dir ?? '.';
	if (/^(?:[a-z]:)?[\\/]/i.test(value) || value.split(/[\\/]/).includes('..'))
		throw new Error('outputDir must be a relative path inside the root');
	return value.replace(/\\/g, '/');
}

const { filePath: _validateFile, ...validateInput } = schemas.ValidatePresentationSchema.shape;
const { filePath: _repairFile, ...repairInput } = schemas.RepairPresentationSchema.shape;

/** PowerPoint tools over the package itself: Markdown conversion, validation and repair. */
export const packageTools: readonly AutomationTool[] = [
	deckTool(
		'read',
		'convert_to_markdown',
		'Convert the presentation to Markdown format.',
		schemas.ConvertToMarkdownSchema,
		(ctx, params) =>
			conversionTools.convertToMarkdown(ctx, {
				...params,
				outputDir: relativeDir(params.outputDir),
			}),
	),
	readTool({
		name: 'validate_presentation',
		description: 'Validate PPTX structural integrity (content types, relationships, XML)',
		format: 'pptx',
		extensions: DECK_READ_EXTENSIONS,
		input: validateInput,
		run: (bytes) => validationTools.validatePresentation(arrayBufferOf(bytes)),
	}),
	editTool({
		name: 'repair_presentation',
		description: 'Auto-repair common PPTX structural issues',
		format: 'pptx',
		extensions: DECK_EDIT_EXTENSIONS,
		input: repairInput,
		run: async (bytes) => {
			const repaired = await validationTools.repairPresentation(arrayBufferOf(bytes));
			return repaired.result.repairCount > 0
				? { ...repaired.result, bytes: new Uint8Array(repaired.repairedBytes) }
				: { ...repaired.result, bytes: undefined };
		},
	}),
];
