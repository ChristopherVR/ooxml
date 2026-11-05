/**
 * Where PowerPoint really puts an element, for the device-pixel stroke snap.
 *
 * The model keeps `x`, `y`, `width` and `height` in whole slide pixels, but
 * PowerPoint rounds every edge to a device pixel from the file's EMU, so a
 * stroke can land one pixel away whenever those whole pixels differ from the
 * real position (a 50.5 pt tall rectangle is 67.33 px, not 67). The parser
 * keeps the exact EMU next to each value (`xEmu`, `yEmu`, `widthEmu`,
 * `heightEmu`), so the snap can use them without a change to the model.
 *
 * PowerPoint also quantises the geometry it paints to its master unit, 1/576
 * inch (1587.5 EMU, an eighth of a point): a rectangle at 290.4 pt sits at
 * 290.375 pt (measured at 1280, 1366 and 1920 px wide). {@link exactDeviceBox}
 * applies the same rounding.
 *
 * @module render/device-pixel-exact
 */

const EMU_PER_PX = 9525;
/** PowerPoint's master unit (1/576 inch) in EMU. */
const EMU_PER_MASTER_UNIT = 1587.5;
/** Master units per slide pixel (a pixel is 1/96 inch). */
const MASTER_UNITS_PER_PX = EMU_PER_PX / EMU_PER_MASTER_UNIT;

/** A box in slide pixels. */
export interface DeviceBox {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** The element fields {@link exactDeviceBox} reads. */
export interface ExactBoxSource extends DeviceBox {
	xEmu?: number | undefined;
	yEmu?: number | undefined;
	widthEmu?: number | undefined;
	heightEmu?: number | undefined;
}

/**
 * `model` slide pixels, or the master-unit rounded position of `emu` when it
 * is still the value the model was parsed from: an element that was moved or
 * resized since load has a model value that no longer rounds to its EMU, and
 * the model wins.
 */
function exactPx(model: number, emu: number | undefined): number {
	if (typeof emu !== 'number' || !Number.isFinite(emu)) {
		return model;
	}
	return Math.round(emu / EMU_PER_PX) === Math.round(model)
		? Math.round(emu / EMU_PER_MASTER_UNIT) / MASTER_UNITS_PER_PX
		: model;
}

/** The box PowerPoint positions `element` at, in slide pixels (the model's where unknown). */
export function exactDeviceBox(element: ExactBoxSource): DeviceBox {
	return {
		x: exactPx(element.x, element.xEmu),
		y: exactPx(element.y, element.yEmu),
		width: exactPx(element.width, element.widthEmu),
		height: exactPx(element.height, element.heightEmu),
	};
}
