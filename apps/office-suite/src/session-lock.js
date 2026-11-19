/**
 * Holding a document still while something writes it (a save, an export, an AI change).
 *
 * The lock refuses input rather than making the editor `inert`: `inert` also takes keyboard
 * focus away, which sent every keyboard and screen reader user back to the top of the page each
 * time they saved. Focus and the selection stay where they are; keys and pointer presses are
 * refused until the write is done. Tab still moves focus, so nobody is held in place.
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

/** Refuse input inside `node` until the returned function runs. Locks may overlap. */
export function lockInput(node) {
	const count = held.get(node) ?? 0;
	held.set(node, count + 1);
	if (count === 0) {
		for (const type of REFUSED) node.addEventListener(type, refuse, CAPTURE);
		node.setAttribute('aria-busy', 'true');
	}
	let released = false;
	return () => {
		if (released) return;
		released = true;
		const left = held.get(node) - 1;
		held.set(node, left);
		if (left > 0) return;
		for (const type of REFUSED) node.removeEventListener(type, refuse, CAPTURE);
		node.removeAttribute('aria-busy');
	};
}
