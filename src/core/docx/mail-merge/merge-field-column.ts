// The column a MERGEFIELD names, kept apart from the merge engine so page layout can read it
// without loading the mail merge.
import { parseFieldInstruction } from '../field-format';

/** The switches of MERGEFIELD: `\b` and `\f` text before and after, `\m` mapped, `\v` vertical. */
export const MERGE_SWITCHES = { b: 'arg', f: 'arg', m: 'flag', v: 'flag' } as const;

/** The column a MERGEFIELD instruction names (quoted names keep their spaces). */
export function mergeFieldColumn(instruction: string): string | undefined {
	const field = parseFieldInstruction(instruction, MERGE_SWITCHES, { numberPictures: true });
	return field?.name === 'MERGEFIELD' && field.args.length === 1 ? field.args[0] : undefined;
}
