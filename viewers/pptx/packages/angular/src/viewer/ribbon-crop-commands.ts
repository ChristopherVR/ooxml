/**
 * ribbon-crop-commands.ts: what a ribbon Crop split button does, shared by
 * both crop entry points (Home > Arrange and Picture Format > Size). The main
 * half toggles crop mode through the crop service; a menu pick (an aspect
 * ratio, Fill or Fit) commits any live session and crops the selected picture
 * in one undoable update.
 */
import { computed } from '@angular/core';
import type { Signal } from '@angular/core';
import {
	canCropElement,
	cropFill,
	cropFit,
	cropToAspectRatio,
	dispatchCropRequest,
	getCachedNativeImageSize,
	getImageSrc,
} from 'ooxml-ui/pptx';
import type { CropElementUpdate } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { EditorStateService } from './editor-state.service';
import type { LoadContentService } from './load-content.service';
import type { PictureCropService } from './picture-crop.service';

export interface RibbonCropDeps {
	editor: EditorStateService;
	crop: PictureCropService | null;
	loader: LoadContentService | null;
	slideIndex: () => number;
	selectedElement: () => PptxElement | null;
	canEdit: () => boolean;
}

export interface RibbonCropCommands {
	/** The single selected picture, when it can be cropped. */
	target: Signal<PptxElement | null>;
	/** Crop mode is open on that picture (the button is pressed). */
	active(): boolean;
	/** Run a Crop split button request (`value` is a menu pick, or none for the toggle). */
	run(value: string | number | undefined): void;
}

export function createRibbonCropCommands(deps: RibbonCropDeps): RibbonCropCommands {
	const target = computed<PptxElement | null>(() => {
		const el = deps.selectedElement();
		return deps.editor.selectedIds().length === 1 && canCropElement(el) ? el : null;
	});

	/** Commit any live crop session, then apply `compute` as ONE undoable update. */
	const applyOneClick = (compute: (el: PptxElement) => CropElementUpdate): void => {
		deps.crop?.commit();
		const id = target()?.id;
		const el = id
			? deps.editor.slides()[deps.slideIndex()]?.elements.find((item) => item.id === id)
			: undefined;
		if (!el || !deps.canEdit()) {
			return;
		}
		deps.editor.updateElement(deps.slideIndex(), el.id, compute(el) as Partial<PptxElement>);
	};

	return {
		target,
		active: () => deps.crop?.isCropping(target()?.id ?? null) ?? false,
		run(value) {
			const natural = (el: PptxElement) =>
				getCachedNativeImageSize(
					getImageSrc(el, deps.loader?.mediaDataUrls() ?? new Map<string, string>()),
				);
			dispatchCropRequest(value, {
				toggle: () => deps.crop?.toggle(deps.slideIndex(), target()),
				aspect: (width, height) => applyOneClick((el) => cropToAspectRatio(el, width, height)),
				fill: () => applyOneClick((el) => cropFill(el, natural(el))),
				fit: () => applyOneClick((el) => cropFit(el, natural(el))),
			});
		},
	};
}
