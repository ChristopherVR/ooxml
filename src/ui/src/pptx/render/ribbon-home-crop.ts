/**
 * The picture Crop split button, defined once for both places the ribbon shows
 * it: Home > Arrange (`home.arrange.crop`) and Picture Format > Size
 * (`pictureFormat.size.crop`). Both draw the same menu (the aspect-ratio
 * presets, then Fill and Fit), carry the same crop hooks and read one crop
 * state, so the two buttons can never disagree.
 *
 * @module render/ribbon-home-crop
 */
import { control, text } from './ribbon-home-family-helpers';
import { HOME_CROP_ITEMS, parseCropValue } from './ribbon-home-menus';
import type {
	RibbonHomeControlSpec,
	RibbonHomeControlState,
	RibbonHomeViewState,
} from './ribbon-home-spec';

/** The control ids that are the picture Crop split button. */
export type RibbonCropControlId = 'home.arrange.crop' | 'pictureFormat.size.crop';

/**
 * Hooks every crop entry point carries: the bindings keep crop mode running
 * when a press lands on `data-pptx-ribbon-control="crop"` / `"crop-menu"`.
 */
const CROP_HOOK = { 'data-pptx-ribbon-control': 'crop' } as const;
const CROP_MENU_HOOK = { 'data-pptx-ribbon-control': 'crop-menu' } as const;

/**
 * The Crop split button: the main half toggles crop mode, the caret opens the
 * Crop to Aspect Ratio menu. `large` draws Office's tall Size-group button
 * (glyph over the caption); the Home strip keeps its compact one.
 */
export function cropControlSpec(id: RibbonCropControlId, large = false): RibbonHomeControlSpec {
	const caret = { labelKey: 'pptx.image.cropToAspectRatio', fallback: 'Crop to Aspect Ratio' };
	if (large) {
		return control(id, 'pptx.image.crop', 'Crop', undefined, {
			kind: 'menu',
			items: HOME_CROP_ITEMS,
			hintKey: 'pptx.image.cropHint',
			icon: 'home.arrange.crop',
			text: text('pptx.image.crop', 'Crop'),
			large: true,
			attrs: CROP_HOOK,
			caret: { ...caret, attrs: CROP_MENU_HOOK },
		});
	}
	return control(id, 'pptx.image.crop', 'Crop', undefined, {
		kind: 'menu',
		popup: true,
		items: HOME_CROP_ITEMS,
		hintKey: 'pptx.image.cropHint',
		attrs: { ...CROP_HOOK, 'data-pptx-chrome': 'crop-main' },
		caret: { ...caret, attrs: { ...CROP_MENU_HOOK, 'data-pptx-chrome': 'crop-caret' } },
	});
}

export interface CropHomeInput {
	editable: boolean;
	/** A single croppable picture is selected; an open crop session also enables the button. */
	canCrop: boolean;
	/** On-canvas crop mode is active (the Crop button is pressed). */
	cropActive: boolean;
	/** The host hid the crop entry points (the `crop` toolbar action). */
	hideCrop?: boolean;
}

/** The state of one Crop split button: its main half and its caret. */
export function cropHomeControls(
	id: RibbonCropControlId,
	input: CropHomeInput,
): RibbonHomeViewState['controls'] {
	const crop: RibbonHomeControlState = {
		disabled: !input.editable || !(input.canCrop || input.cropActive),
		hidden: input.hideCrop,
	};
	return {
		[id]: { ...crop, pressed: input.cropActive },
		[`${id}#caret`]: crop,
	} as RibbonHomeViewState['controls'];
}

/** Picture Format > Size: the Crop split button. */
export function pictureSizeHomeControls(input: CropHomeInput): RibbonHomeViewState['controls'] {
	return cropHomeControls('pictureFormat.size.crop', input);
}

/** What a binding does for each Crop split button request. */
export interface CropRequestHandlers {
	/** The main half: enter crop mode, or commit it when already open. */
	toggle(): void;
	aspect(ratioWidth: number, ratioHeight: number): void;
	fill(): void;
	fit(): void;
}

/**
 * Route a Crop split button request (`value` is the picked menu row, or none
 * for the main half) to the binding's crop commands. Every binding and both
 * crop entry points dispatch through this one decision.
 */
export function dispatchCropRequest(
	value: string | number | undefined,
	handlers: CropRequestHandlers,
): void {
	if (value === undefined) {
		handlers.toggle();
		return;
	}
	const action = parseCropValue(value);
	if (action?.kind === 'aspect') {
		handlers.aspect(action.width, action.height);
	} else if (action?.kind === 'fill') {
		handlers.fill();
	} else if (action?.kind === 'fit') {
		handlers.fit();
	}
}
