/**
 * The keys of the speaker-notes field, one rule for every binding.
 *
 * The field keeps typing to itself because the editor keymap ignores text
 * fields; it must not stop keys from bubbling, or F6 never reaches the pane
 * cycle and the keyboard is trapped. Tab, Shift+Tab and F6 move on as they do
 * anywhere else. Escape goes back to the slide, and Ctrl/Cmd+Enter commits the
 * notes and goes back to the slide.
 *
 * @module render/notes/notes-keys
 */
import { focusSlidePane } from '../slide-keyboard-focus';

/**
 * Handle a keydown on the notes field. `commit` saves the notes now; a field that
 * saves on blur can omit it, since going back to the slide blurs it.
 */
export function handleNotesFieldKey(event: KeyboardEvent, commit?: () => void): void {
	const mod = event.ctrlKey || event.metaKey;
	if (event.isComposing) {
		return;
	}
	if (event.key === 'Enter' && mod && !event.altKey) {
		event.preventDefault();
		commit?.();
	} else if (event.key !== 'Escape' || mod || event.altKey || event.shiftKey) {
		return;
	}
	focusSlidePane(event.target);
}
