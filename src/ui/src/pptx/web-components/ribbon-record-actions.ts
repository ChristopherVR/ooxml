import {
	clearSlideTimings,
	containsCameraClip,
	removeCameraClips,
} from 'ooxml-core/pptx/editor/render/slide-operations';
import { showOfficeServiceDialog } from '../../dialog/office-service-dialog';
import type { Translator } from '../i18n/translator';
import { generateElementId } from '../render/element-clipboard';
import type { RibbonControlId } from '../render/customization';
import { slideHasCameo } from '../render/ribbon-insert-state';
import {
	recordScopeAvailable,
	type RecordScope,
	type RecordTarget,
} from '../render/ribbon-record-commands';
import { createCameoElement } from './cameo-insert';
import { recordCameraClip } from './record-camera-clip';

/** Handled by the shared Record tab from every binding. */
export function runRecordRibbonAction(
	host: HTMLElement,
	id: RibbonControlId,
	target: () => RecordTarget | undefined,
	t: Translator,
): void {
	if (id !== 'record.camera.cameo' && id !== 'record.manage.clear' && id !== 'record.manage.reset')
		return;
	const initial = target();
	if (!initial) return;
	const action = id;
	const scopes = (['current', 'all'] as const).filter((scope) =>
		recordScopeAvailable(action, initial, scope),
	);
	if (host.querySelector('office-ui-dialog[open]')) return;
	const doc = host.ownerDocument;
	const cameo = action === 'record.camera.cameo';
	if (cameo ? !initial.slides[initial.slideIndex] : !scopes.length) return;
	const reset = action === 'record.manage.reset';
	const heading = t(
		cameo ? 'pptx.record.cameo' : reset ? 'pptx.record.resetToCameo' : 'pptx.record.clearTimings',
	);
	const dialog = showOfficeServiceDialog(host, heading, t('pptx.common.close'));
	let busy = false;
	const liveTarget = (scope: RecordScope) => {
		const live = target();
		return dialog.open &&
			host.isConnected &&
			live?.slides === initial.slides &&
			(scope === 'all' || live.slideIndex === initial.slideIndex) &&
			recordScopeAvailable(action, live, scope)
			? live
			: undefined;
	};
	for (const scope of scopes) {
		const button = doc.createElement('button');
		button.type = 'button';
		button.textContent = t(
			scope === 'current' ? 'pptx.print.currentSlide' : 'pptx.print.allSlides',
		);
		button.addEventListener('click', async () => {
			if (busy) return;
			busy = true;
			try {
				const element =
					cameo || reset
						? await createCameoElement(doc, initial.canvasSize, t('pptx.record.cameo'))
						: undefined;
				const live = liveTarget(scope);
				if (!live) {
					dialog.close();
					return;
				}
				if (reset && element) {
					live.commitSlides?.(
						live.slides.map((slide, index) => {
							if (
								(scope === 'current' && index !== live.slideIndex) ||
								!containsCameraClip(slide.elements)
							)
								return slide;
							const elements = removeCameraClips(slide.elements);
							if (!slideHasCameo(slide)) elements.push({ ...element, id: generateElementId() });
							return { ...slide, elements, isDirty: true };
						}),
					);
				} else if (element && scope === 'current') live.insert(element);
				else if (element) {
					live.commitSlides?.(
						live.slides.map((slide) =>
							slideHasCameo(slide)
								? slide
								: {
										...slide,
										isDirty: true,
										elements: [...slide.elements, { ...element, id: generateElementId() }],
									},
						),
					);
				} else
					live.commitSlides?.(
						clearSlideTimings(live.slides, scope === 'current' ? live.slideIndex : undefined),
					);
			} catch {
				const error = doc.createElement('p');
				error.setAttribute('role', 'alert');
				error.textContent = t('pptx.graphic.insertFailed');
				dialog.append(error);
				return;
			} finally {
				busy = false;
			}
			dialog.close();
		});
		dialog.append(button, ' ');
	}
	if (cameo) {
		const record = doc.createElement('button');
		record.type = 'button';
		record.textContent = t('pptx.record.cameraClip');
		record.addEventListener('click', () => {
			if (busy) return;
			const live = target();
			dialog.close();
			if (live?.slides === initial.slides && live.slideIndex === initial.slideIndex) {
				recordCameraClip(host, live, target, t);
			}
		});
		dialog.append(record);
	}
	dialog.show();
}
