import type { ShapeStyle, XmlObject } from '../../types';
import { clearFillChoiceGroup } from './fill-choice-group';

const isOn = (value: unknown): boolean => value === '1' || value === 'true' || value === true;

/**
 * Write `<p:sp useBgFill>` (ECMA-376 §19.3.1.43) from `style.useBackgroundFill`.
 *
 * The fill writer leaves a background-filled shape's `spPr` alone (its fill
 * fields hold the slide background the load resolved), so the flag itself is
 * what an edit changes. Turning it on drops the shape's own fill child, which
 * the flag replaces; turning an authored flag off removes the attribute and
 * lets the fill writer emit the fill the user chose. Only `p:sp` (CT_Shape)
 * carries the attribute; connectors and other nodes are left untouched, as is
 * an attribute that is already off.
 */
export function writeBackgroundFillFlag(
	shape: XmlObject,
	spPr: XmlObject,
	style: ShapeStyle,
): void {
	if (shape['p:nvSpPr'] === undefined) {
		return;
	}
	const authored = isOn(shape['@_useBgFill']);
	if (style.useBackgroundFill) {
		if (!authored) {
			shape['@_useBgFill'] = '1';
			clearFillChoiceGroup(spPr);
		}
		return;
	}
	if (authored) {
		delete shape['@_useBgFill'];
	}
}
