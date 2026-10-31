import { attribute, children, VISIO_NS } from './sheet';
import { fail } from './package-common';
import { shapeIdOf, topLevelShapes } from './edit-container-members';

const guid = (): string => `{${crypto.randomUUID().toUpperCase()}}`;

/** The highest shape ID anywhere on the page, sub-shapes included. */
function highestId(root: Element): number {
	let max = 0;
	for (const shape of Array.from(root.getElementsByTagName('Shape')))
		max = Math.max(max, Number(attribute(shape, 'ID')) || 0);
	return max;
}

/** The sub-shape of a lane that carries its heading text (the master's text shape). */
export function headingShape(lane: Element): Element | undefined {
	const parts = Array.from(lane.getElementsByTagName('Shape')).filter((shape) => shape !== lane);
	return (
		parts.find((shape) =>
			children(shape, 'Section').some(
				(section) =>
					attribute(section, 'N') === 'User' &&
					children(section, 'Row').some((row) => attribute(row, 'N') === 'HeadingPos'),
			),
		) ?? parts[parts.length - 1]
	);
}

/**
 * A new lane or phase separator for a chart Visio built: a copy of the one it goes beside, with
 * its own IDs and GUIDs, so it is an instance of the same master with the same links to its
 * list and the frame. Placed right after the reference; the caller sizes and orders it. A
 * `heading` of undefined leaves the heading without text of its own, as Visio's new separators
 * are (the master's "Phase" shows).
 */
export function cloneNativeLane(
	root: Element,
	reference: Element,
	id: string,
	heading: string | undefined,
): Element {
	if (topLevelShapes(root).some((shape) => shapeIdOf(shape) === id))
		fail('INVALID_EDIT', 'The new lane ID is already used.');
	const copy = reference.cloneNode(true) as Element;
	let next = Math.max(highestId(root), Number(id));
	copy.setAttribute('ID', id);
	copy.setAttribute('UniqueID', guid());
	copy.removeAttribute('NameU');
	copy.removeAttribute('Name');
	for (const part of Array.from(copy.getElementsByTagName('Shape')))
		if (part !== copy) {
			part.setAttribute('ID', String(++next));
			if (part.hasAttribute('UniqueID')) part.setAttribute('UniqueID', guid());
		}
	const target = headingShape(copy);
	if (target) {
		for (const text of children(target, 'Text')) target.removeChild(text);
		if (heading !== undefined) {
			const text = target.ownerDocument.createElementNS(VISIO_NS, 'Text');
			text.textContent = heading;
			const sub = children(target, 'Shapes')[0];
			target.insertBefore(text, sub ?? null);
		}
	}
	reference.parentNode!.insertBefore(copy, reference.nextSibling);
	return copy;
}
