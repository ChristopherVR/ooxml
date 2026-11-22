import type { ZoomPptxElement } from 'ooxml-core/pptx';
import { SvgExporter } from 'ooxml-core/pptx/converter';
import { prepareSummaryZoom } from 'ooxml-core/pptx/editor/render/summary-zoom-operations';
import type { OfficeUiCheckbox } from '../../form/checkbox';
import type { OfficeUiRadio } from '../../form/radio';
import { renderLocalSvg } from '../../form/local-svg';
import { createRibbonZoom, zoomChoices } from '../render/ribbon-zoom-insert';
import { generateElementId } from '../render/element-clipboard';
import type { RibbonGraphicTarget, RibbonZoom } from '../render/ribbon-insert-state';
import type { InsertText } from './ribbon-graphic-insert';
import { createInsertDialogFrame } from './insert-dialog-frame';
import { definePptxCheckbox, definePptxRadio } from './office-aliases';

export function pickRibbonZoom(
	host: HTMLElement,
	target: RibbonGraphicTarget,
	kind: RibbonZoom,
	text: InsertText,
	current: () => RibbonGraphicTarget | undefined,
	onInserted?: () => void,
) {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	if (registry) {
		definePptxCheckbox(registry);
		definePptxRadio(registry);
	}
	const heading = {
		slide: text('pptx.zoom.slideZoom', 'Slide Zoom'),
		section: text('pptx.zoom.sectionZoom', 'Section Zoom'),
		summary: text('pptx.zoom.summaryZoom', 'Summary Zoom'),
	}[kind];
	const options = zoomChoices(target.slides, target.slideIndex, kind);
	const body = doc.createElement('div');
	const list = doc.createElement('fieldset');
	const title = doc.createElement('legend');
	const multiple = kind === 'summary';
	title.textContent = multiple
		? text('pptx.graphic.zoomSummaryTargets', 'Choose the first slide of each section')
		: kind === 'slide'
			? text('pptx.graphic.zoomSlideTargets', 'Choose a slide')
			: text('pptx.graphic.zoomSectionTargets', 'Choose sections');
	list.append(title);
	body.append(list);
	const keepUnused = doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox;
	keepUnused.checked = true;
	if (multiple && target.sections.length) {
		const keepLabel = doc.createElement('label');
		const name = text('pptx.graphic.zoomKeepUnusedSections', 'Keep unused sections');
		keepUnused.setAttribute('aria-label', name);
		keepLabel.append(keepUnused, ` ${name}`);
		body.append(keepLabel);
	}
	const settings = zoomInsertSettings(body, kind, text);
	const sectionStarts = new Set<string>();
	const radioGroup = generateElementId();
	const entries = options.map((choice) => {
		const label = doc.createElement('label');
		label.style.display = 'block';
		const input = multiple
			? (doc.createElement('pptx-ui-checkbox') as OfficeUiCheckbox)
			: (doc.createElement('pptx-ui-radio') as OfficeUiRadio);
		if ('name' in input) input.name = radioGroup;
		const startsSection = Boolean(
			choice.slide.sectionId && !sectionStarts.has(choice.slide.sectionId),
		);
		if (choice.slide.sectionId) sectionStarts.add(choice.slide.sectionId);
		input.checked = multiple && startsSection;
		const slideLabel = text('pptx.zoom.slideNumber', 'Slide {{number}}', {
			number: choice.index + 1,
		});
		const name =
			kind !== 'section'
				? [slideLabel, choice.slide.name || (startsSection && choice.slide.sectionName)]
						.filter(Boolean)
						.join(' · ')
				: choice.slide.sectionName || slideLabel;
		input.setAttribute('aria-label', name);
		label.append(input, ` ${name}`);
		list.append(label);
		return { input, choice, name };
	});
	return new Promise<ReturnType<typeof createRibbonZoom>>((resolve, reject) => {
		let committing = false;
		const frame = createInsertDialogFrame(
			host,
			body,
			() => {
				const selected = entries.filter(({ input }) => input.checked);
				if (committing || !selected.length || !settings.duration.checkValidity()) return;
				committing = true;
				for (const { input } of entries) input.disabled = true;
				keepUnused.disabled = true;
				const { cover, ...format } = settings.read();
				settings.disable();
				sync();
				void Promise.all(
					selected.map(async ({ choice, name }) => {
						const { width, height } = target.canvasSize;
						const ratio = Math.min(1, 360 / Math.max(width, height));
						if (cover) {
							return { ...choice, title: name, preview: await zoomCover(doc, cover) };
						}
						const svg = SvgExporter.exportSlide(choice.slide, width, height, {
							transparentBackground: format.zoomShowBackground === false,
						});
						const preview = await renderLocalSvg(
							doc,
							svg,
							Math.max(1, Math.round(width * ratio)),
							Math.max(1, Math.round(height * ratio)),
						);
						return { ...choice, title: name, preview: preview.toDataURL('image/png') };
					}),
				)
					.then((previews) => {
						if (!frame.dialog.open) return;
						const live = current();
						if (!live) return frame.dialog.close();
						if (multiple) {
							const prepared = prepareSummaryZoom(
								live.slides,
								live.sections,
								previews.map(({ slide, title, preview }) => ({
									slideId: slide.id,
									title,
									preview,
								})),
								{
									name: heading,
									defaultSectionName: text('pptx.sections.defaultName', 'Untitled Section'),
									keepUnused: keepUnused.checked,
								},
							);
							if (prepared) {
								const zoom = createRibbonZoom(
									'summary',
									prepared.entries,
									live.canvasSize,
									heading,
								);
								if (zoom && live.insertSummary) {
									Object.assign(zoom, format);
									for (const tile of zoom.summaryTargets ?? []) Object.assign(tile, format);
									const { width, height } = live.canvasSize;
									prepared.result.slides[prepared.result.slideIndex]?.elements.push(
										{
											type: 'text',
											id: generateElementId(),
											text: heading,
											x: width * 0.09,
											y: height * 0.035,
											width: width * 0.82,
											height: height * 0.1,
											textStyle: {
												fontSize: 32,
												bold: true,
												color: live.themeColors?.dk1 ?? '#000000',
											},
										},
										zoom,
									);
									live.insertSummary(prepared.result);
									frame.dialog.close();
									onInserted?.();
									return;
								}
							}
							frame.dialog.close();
							return;
						}
						const choices = zoomChoices(live.slides, live.slideIndex, kind);
						const resolved = previews.flatMap((preview) => {
							const choice = choices.find(
								({ slide, sectionId }) =>
									slide.id === preview.slide.id && sectionId === preview.sectionId,
							);
							return choice ? [{ ...preview, ...choice }] : [];
						});
						if (resolved.length === previews.length) {
							const zoom = createRibbonZoom(kind, resolved, live.canvasSize, heading);
							resolve(zoom && { ...zoom, ...format });
						}
						frame.dialog.close();
					})
					.catch((error: unknown) => {
						if (!frame.dialog.open) return;
						reject(error);
						frame.dialog.close();
					});
			},
			() => {
				resolve(undefined);
				frame.dialog.remove();
			},
		);
		const sync = () =>
			frame.sync(
				heading,
				text('pptx.common.close', 'Close'),
				text('pptx.common.cancel', 'Cancel'),
				text('pptx.common.insert', 'Insert'),
				!committing &&
					settings.duration.checkValidity() &&
					entries.some(({ input }) => input.checked),
			);
		list.addEventListener('change', sync);
		settings.duration.addEventListener('input', sync);
		sync();
		frame.dialog.show();
	});
}

function zoomInsertSettings(body: HTMLElement, kind: RibbonZoom, text: InsertText) {
	const doc = body.ownerDocument;
	const options = doc.createElement('fieldset');
	const heading = doc.createElement('legend');
	heading.textContent = text('pptx.graphic.zoomOptions', 'Zoom Options');
	options.append(heading);
	const input = (type: string, label: string) => {
		const control = doc.createElement('input');
		control.type = type;
		const name = doc.createElement('label');
		name.style.display = 'block';
		name.append(control, ` ${label}`);
		options.append(name);
		return control;
	};
	const back = input('checkbox', text('pptx.graphic.zoomReturn', 'Return to Zoom'));
	back.checked = kind !== 'slide';
	const transparent = input(
		'checkbox',
		text('pptx.graphic.zoomTransparent', 'Transparent background'),
	);
	const duration = input('number', text('pptx.transition.duration', 'Duration (ms)'));
	duration.min = '0';
	duration.max = '60000';
	duration.step = '1';
	duration.style.width = '9ch';
	const thumbnail =
		kind === 'summary'
			? undefined
			: input('file', text('pptx.graphic.zoomCoverImage', 'Cover image'));
	if (thumbnail) {
		thumbnail.accept = 'image/png,image/jpeg,image/webp,image/gif';
		thumbnail.style.maxWidth = '100%';
	}
	body.append(options);
	return {
		duration,
		disable: () => {
			options.disabled = true;
		},
		read: (): Pick<
			ZoomPptxElement,
			'returnToParent' | 'transitionDurationMs' | 'zoomShowBackground' | 'zoomImageType'
		> & { cover?: File } => {
			const cover = thumbnail?.files?.[0];
			return {
				returnToParent: back.checked,
				transitionDurationMs: duration.value ? duration.valueAsNumber : undefined,
				zoomImageType: cover ? 'cover' : 'preview',
				zoomShowBackground: !transparent.checked,
				...(cover ? { cover } : {}),
			};
		},
	};
}

async function zoomCover(doc: Document, file: File): Promise<string> {
	if (file.size > 16 * 1024 * 1024) throw new Error('The Zoom cover image is too large.');
	const bitmap = await createImageBitmap(file);
	try {
		const scale = Math.min(1, 360 / Math.max(bitmap.width, bitmap.height));
		const canvas = doc.createElement('canvas');
		canvas.width = Math.max(1, Math.round(bitmap.width * scale));
		canvas.height = Math.max(1, Math.round(bitmap.height * scale));
		const context = canvas.getContext('2d');
		if (!context) throw new Error('Could not render the Zoom cover image.');
		context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
		return canvas.toDataURL('image/png');
	} finally {
		bitmap.close();
	}
}
