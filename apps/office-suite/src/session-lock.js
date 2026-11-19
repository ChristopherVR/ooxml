/**
 * Holding a document still while something writes it (a save, an export, an AI change).
 *
 * The lock refuses input rather than making the editor `inert`: `inert` also takes keyboard
 * focus away, which sent every keyboard and screen reader user back to the top of the page each
 * time they saved. Focus and the selection stay where they are; keys and pointer presses are
 * refused until the write is done. Tab still moves focus, so nobody is held in place.
 *
 * An input method's composition cannot be refused: the browser writes it into the editor whatever
 * the page does. The lock records the text each composition commits instead (`takeComposed`), so
 * the session can type it again when a reload replaced the editor's content, or count it as an
 * unsaved edit when the content stayed.
 */
const REFUSED = [
	'beforeinput',
	'keydown',
	'keypress',
	'paste',
	'cut',
	'drop',
	'dragstart',
	'pointerdown',
	'mousedown',
	'click',
	'dblclick',
	'contextmenu',
	'touchstart',
];

function refuse(event) {
	event.stopPropagation();
	if (event.type === 'keydown' && event.key === 'Tab') return;
	event.preventDefault();
}

const CAPTURE = { capture: true };
const held = new WeakMap();
const composed = new WeakMap();

function record(event) {
	if (event.data)
		composed.set(event.currentTarget, (composed.get(event.currentTarget) ?? '') + event.data);
}

/**
 * Refuse input inside `node` until the returned function runs. Locks may overlap; the function
 * returns true when it released the last one.
 */
export function lockInput(node) {
	const count = held.get(node) ?? 0;
	held.set(node, count + 1);
	if (count === 0) {
		for (const type of REFUSED) node.addEventListener(type, refuse, CAPTURE);
		node.addEventListener('compositionend', record, CAPTURE);
		node.setAttribute('aria-busy', 'true');
	}
	let released = false;
	return () => {
		if (released) return false;
		released = true;
		const left = held.get(node) - 1;
		held.set(node, left);
		if (left > 0) return false;
		for (const type of REFUSED) node.removeEventListener(type, refuse, CAPTURE);
		node.removeEventListener('compositionend', record, CAPTURE);
		node.removeAttribute('aria-busy');
		return true;
	};
}

/** Whether a lock holds `node`. */
export const isLocked = (node) => (held.get(node) ?? 0) > 0;

/** The text input methods committed inside `node` while it was locked, since the last call. */
export function takeComposed(node) {
	const text = composed.get(node) ?? '';
	composed.delete(node);
	return text;
}
