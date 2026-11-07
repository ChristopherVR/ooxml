/**
 * Child order of a picture's `a:blip` on save.
 *
 * @module runtime/blip-extension-order
 */

import type { XmlObject } from '../../types';

/**
 * `CT_Blip` is any number of effects followed by an optional `a:extLst`. The
 * picture effect writers (`applyImageEffectsToBlip`) add an effect the blip did not have as a new last child, so
 * on a blip that already had an extension list (PowerPoint writes
 * `a14:useLocalDpi` on most pictures) the effect landed after it. Moving the
 * list back to the end keeps the schema order.
 */
export function keepExtensionListLast(blip: XmlObject): void {
	const key = Object.keys(blip).find((candidate) => candidate.split(':').at(-1) === 'extLst');
	if (key === undefined) {
		return;
	}
	const value = blip[key];
	delete blip[key];
	blip[key] = value;
}
