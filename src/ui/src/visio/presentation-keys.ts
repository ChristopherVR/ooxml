/** A Presentation Mode command produced by one key press. */
export type PresentationKeyAction =
	| { type: 'next' }
	| { type: 'previous' }
	| { type: 'first' }
	| { type: 'last' }
	/** Jump to a 1-based foreground page number typed before Enter. */
	| { type: 'goto'; page: number }
	| { type: 'exit' }
	/** A digit was added to the pending page number. */
	| { type: 'buffering'; digits: string }
	| { type: 'none' };

/** Digits typed so far for "type a page number, then Enter". One per running presentation. */
export interface PresentationKeyBuffer {
	digits: string;
}

const NEXT = new Set(['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Spacebar', 'Enter']);
const PREVIOUS = new Set(['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace']);

/**
 * Visio's Presentation Mode keys: Right, Down, Page Down, Space and Enter advance; Left, Up,
 * Page Up and Backspace go back; Home and End jump to the first and last page; Escape exits;
 * digits then Enter jump to that page. Chords are left alone. The PowerPoint slide-show map
 * (`pptx/render/presentation-keymap`) adds pens, screens and menus Visio does not have, and
 * lives in the separately compiled pptx project, so Visio keeps this small map.
 */
export function mapPresentationKey(
	event: { key: string; ctrlKey?: boolean; metaKey?: boolean; altKey?: boolean },
	buffer: PresentationKeyBuffer,
): PresentationKeyAction {
	const { key } = event;
	if (event.ctrlKey || event.metaKey || event.altKey) return { type: 'none' };
	if (key.length === 1 && key >= '0' && key <= '9') {
		buffer.digits = (buffer.digits + key).slice(-4);
		return { type: 'buffering', digits: buffer.digits };
	}
	// A lone modifier (Shift for digits on some layouts) keeps the pending number.
	if (['Shift', 'Control', 'Alt', 'Meta', 'CapsLock'].includes(key)) return { type: 'none' };
	const digits = buffer.digits;
	buffer.digits = '';
	if (key === 'Enter' && digits) {
		const page = Number.parseInt(digits, 10);
		return page > 0 ? { type: 'goto', page } : { type: 'none' };
	}
	if (NEXT.has(key)) return { type: 'next' };
	if (PREVIOUS.has(key)) return { type: 'previous' };
	if (key === 'Home') return { type: 'first' };
	if (key === 'End') return { type: 'last' };
	if (key === 'Escape') return { type: 'exit' };
	// Any other key abandons a pending number, as PowerPoint and Visio do.
	return { type: 'none' };
}
