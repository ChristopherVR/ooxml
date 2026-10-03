import type { EditorContext } from '../context.js';
import { field, invalid, textInput } from './fields.js';
import { showDialog } from './frame.js';

/** Passwords stay in memory for the current workbook and are never persisted in settings. */
export function packagePassword(ctx: EditorContext, saving = false): Promise<string | undefined> {
	const input = textInput(ctx);
	input.type = 'password';
	input.autocomplete = saving ? 'new-password' : 'current-password';
	const again = textInput(ctx);
	again.type = 'password';
	again.autocomplete = 'new-password';
	return showDialog(ctx, {
		name: saving ? 'encrypt-workbook' : 'open-password',
		heading: saving ? 'Encrypt Workbook' : 'Password required',
		body: [
			field(ctx, 'Password:', input),
			...(saving ? [field(ctx, 'Confirm password:', again)] : []),
		],
		opened: () => input.focus(),
		submit: () => {
			if (!input.value) return invalid(ctx, input, 'Enter a password.');
			if (saving && input.value !== again.value)
				return invalid(ctx, again, 'Confirmation password is not identical.');
			return input.value;
		},
	});
}
