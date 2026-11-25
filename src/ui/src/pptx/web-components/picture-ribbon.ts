import { isDirectPicture, resetPicturePatch } from 'ooxml-core/pptx';
import type { PptxImageLikeElement } from 'ooxml-core/pptx';
import { defineElement } from '../../registry';
import { insertDialogText } from '../../dialog/insert-dialog-frame';
import { canInteractWithElement } from '../render/element-locks';
import { definePptxRibbonCommand } from './office-aliases';
import { canEditPicture, canProcessPictureRaster } from './picture-ribbon-state';
import type { PictureRibbonState, PictureStripGroup } from './picture-ribbon-state';

const COMMANDS = {
	'picture-adjust': [
		['adjust.removeBackground', 'removeBackground', 'Remove Background', 'crop'],
		['adjust.compress', 'compress', 'Compress Pictures', 'image'],
		['adjust.changePicture', 'changePicture', 'Change Picture', 'image'],
		['adjust.resetPicture', 'resetPicture', 'Reset Picture', 'reset'],
		['adjust.resetSize', 'resetSize', 'Reset Picture & Size', 'reset'],
	],
	'picture-styles': [
		['pictureStyles.border', 'border', 'Picture Border', 'pencil'],
		['pictureStyles.layout', 'layout', 'Picture Layout', 'layout'],
	],
	'picture-accessibility': [['accessibility.altText', 'altText', 'Alt Text', 'accessibility']],
	'picture-arrange': [
		['arrange.rotateRight', 'rotateRight', 'Rotate Right 90°', 'reset'],
		['arrange.selectionPane', 'selectionPane', 'Selection Pane', 'list'],
	],
} as const;

type PictureCommand = (typeof COMMANDS)[PictureStripGroup][number][1];
export type PictureDialogCommand =
	| Extract<
			PictureCommand,
			'removeBackground' | 'compress' | 'changePicture' | 'resetSize' | 'border'
	  >
	| 'resetPicture';

export interface PictureDialogContext {
	host: HTMLElement;
	picture: PptxImageLikeElement;
	state: () => PictureRibbonState | undefined;
	text: (key: string, fallback: string) => string;
	signal: AbortSignal;
	cancel: () => void;
}

export interface PptxPictureRibbonElement extends HTMLElement {
	family: PictureStripGroup;
	state: PictureRibbonState | undefined;
}

declare global {
	interface HTMLElementTagNameMap {
		'pptx-ui-picture-ribbon': PptxPictureRibbonElement;
	}
}

export function definePptxPictureRibbon(registry: CustomElementRegistry): void {
	if (registry.get('pptx-ui-picture-ribbon')) return;
	definePptxRibbonCommand(registry);
	class PictureRibbon extends HTMLElement implements PptxPictureRibbonElement {
		private controls = new Map<PictureCommand, HTMLElement>();
		private group: PictureStripGroup = 'picture-adjust';
		private model: PictureRibbonState | undefined;
		private active: AbortController | undefined;
		private focused: HTMLElement | null = null;

		set family(value: PictureStripGroup) {
			if (this.group === value) return;
			for (const control of this.controls.values()) control.remove();
			this.controls.clear();
			this.group = value;
			this.sync();
		}
		get family() {
			return this.group;
		}
		set state(value: PictureRibbonState | undefined) {
			if (this.active && !value?.editable) this.active.abort();
			this.model = value;
			this.sync();
		}
		get state() {
			return this.model;
		}

		connectedCallback() {
			this.sync();
		}
		disconnectedCallback() {
			this.active?.abort();
		}

		private sync(): void {
			// The ribbon group treats a full-height inline control as a column beside
			// its stacked dropdown galleries.
			if (this.group === 'picture-adjust' || this.group === 'picture-styles') {
				this.setAttribute('mode', 'inline');
			} else {
				this.removeAttribute('mode');
			}
			this.style.display = 'inline-grid';
			this.style.gridAutoFlow = 'column';
			this.style.gridTemplateRows = 'repeat(3, auto)';
			for (const [id, action, label, icon] of COMMANDS[this.group]) {
				let button = this.controls.get(action);
				if (!button) {
					button = this.ownerDocument.createElement('pptx-ui-ribbon-command');
					button.setAttribute('data-ribbon-control', 'pictureFormat.' + id);
					button.setAttribute('icon', icon);
					button.setAttribute('compact', '');
					button.addEventListener('command-request', () => this.run(action));
					this.controls.set(action, button);
					this.append(button);
				}
				const text = this.label('pptx.pictureFormat.' + action, label);
				const reason = this.reason(action);
				button.setAttribute('label', text);
				button.setAttribute('title', reason ?? text);
				button.toggleAttribute('disabled', Boolean(reason));
			}
		}

		private label(key: string, fallback: string): string {
			return insertDialogText(this.model?.translate, key, fallback);
		}

		private reason(action: PictureCommand): string | undefined {
			if (action === 'layout')
				return this.label(
					'pptx.pictureFormat.layoutUnavailable',
					'Picture layouts cannot be created from a selection yet.',
				);
			const state = this.model;
			const picture = state?.element;
			if (action === 'altText' && isDirectPicture(picture) && state?.altText) return;
			if (action === 'selectionPane' && state?.selectionPane) return;
			if (
				!state?.editable ||
				!isDirectPicture(picture) ||
				!canEditPicture(picture) ||
				action === 'altText' ||
				action === 'selectionPane'
			) {
				return this.label('pptx.pictureFormat.unavailable', 'Select an editable picture.');
			}
			if (this.active) return this.label('pptx.pictureFormat.processing', 'Processing picture…');
			if (
				(action === 'rotateRight' && !canInteractWithElement(picture, 'rotate')) ||
				(action === 'resetSize' && !canInteractWithElement(picture, 'resize')) ||
				((action === 'resetSize' || action === 'resetPicture') &&
					(!canInteractWithElement(picture, 'crop') ||
						(picture.cropShape && !canInteractWithElement(picture, 'changeShapeType'))))
			) {
				return this.label('pptx.pictureFormat.unavailable', 'Select an editable picture.');
			}
			if (
				(action === 'removeBackground' || action === 'compress' || action === 'resetSize') &&
				(!canProcessPictureRaster(picture) ||
					typeof Worker !== 'function' ||
					typeof OffscreenCanvas !== 'function')
			) {
				return this.label(
					'pptx.pictureFormat.rasterOnly',
					'Requires a local PNG or JPEG and browser support for picture processing.',
				);
			}
			if (
				(action === 'removeBackground' || action === 'resetPicture' || action === 'resetSize') &&
				picture.imageEffects?.backgroundRemoval &&
				!picture.originalPictureData &&
				!picture.originalPicturePath
			) {
				return this.label(
					'pptx.pictureFormat.originalUnavailable',
					'The original picture is unavailable.',
				);
			}
		}

		private run(action: PictureCommand): void {
			if (this.reason(action)) return;
			const state = this.model;
			const picture = state?.element;
			if (!state || !isDirectPicture(picture) || action === 'layout') return;
			if (action === 'altText') {
				state.altText?.();
				return;
			}
			if (action === 'selectionPane') {
				state.selectionPane?.();
				return;
			}
			if (action === 'rotateRight') {
				state.patch(picture.id, { rotation: ((picture.rotation ?? 0) + 90) % 360 });
				return;
			}
			if (action === 'resetPicture' && !picture.imageEffects?.backgroundRemoval) {
				state.patch(picture.id, resetPicturePatch(picture));
				return;
			}
			const controller = new AbortController();
			this.active = controller;
			this.focused = this.controls.get(action)?.shadowRoot?.querySelector('button') ?? null;
			controller.signal.addEventListener(
				'abort',
				() => {
					this.active = undefined;
					this.sync();
					const focused = this.focused;
					queueMicrotask(() => focused?.focus());
					this.focused = null;
				},
				{ once: true },
			);
			this.sync();
			const context: PictureDialogContext = {
				host: this,
				picture,
				state: () => this.model,
				text: (key, fallback) => this.label(key, fallback),
				signal: controller.signal,
				cancel: () => controller.abort(),
			};
			void import('./picture-dialogs').then(
				({ runPictureDialog }) => {
					if (!context.signal.aborted) runPictureDialog(action, context);
				},
				() => controller.abort(),
			);
		}
	}
	defineElement(registry, 'pptx-ui-picture-ribbon', PictureRibbon);
}
