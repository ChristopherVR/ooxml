import type { XmlObject, TextStyle } from '../types';
import {
	assignBodyAttributes,
	bodyFlagsFromDrawing,
	bodyFlagsToDrawing,
	objectBodyProperties,
} from './body-properties-neutral';

/**
 * Parse additional boolean body properties from a:bodyPr attributes.
 *
 * Handles the boolean / numeric attributes defined on
 * CT_TextBodyProperties that aren't mapped elsewhere in the runtime:
 *   `compatLnSpc`, `forceAA`, `upright`, `fromWordArt`,
 *   `spcFirstLastPara`, `anchorCtr`, `rtlCol`, and `rot` (degrees).
 * The values are read by the shared DrawingML body codec.
 */
export function parseBodyPrBooleanAttrs(bodyPr: XmlObject, textStyle: TextStyle): void {
	Object.assign(textStyle, bodyFlagsFromDrawing(objectBodyProperties(bodyPr)));
}

/**
 * Write body property boolean attributes to bodyPr XML object.
 *
 * Mirrors {@link parseBodyPrBooleanAttrs}: only emits attributes that the
 * caller explicitly authored. `rtlCol` is intentionally not defaulted here
 * so callers can preserve the loaded value when the model is undefined.
 */
export function writeBodyPrBooleanAttrs(bodyPr: XmlObject, textStyle: TextStyle | undefined): void {
	if (!textStyle) {
		return;
	}
	assignBodyAttributes(bodyPr, bodyFlagsToDrawing(textStyle));
}
