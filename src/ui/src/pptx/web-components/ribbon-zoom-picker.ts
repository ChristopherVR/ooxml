import { SvgExporter } from 'ooxml-core/pptx/converter';
import { prepareSummaryZoom } from 'ooxml-core/pptx/editor/render/summary-zoom-operations';
import { renderLocalSvg } from '../../form/local-svg';
import { createRibbonZoom, zoomChoices } from '../render/ribbon-zoom-insert';
import { generateElementId } from '../render/element-clipboard';
import type { RibbonGraphicTarget, RibbonZoom } from '../render/ribbon-insert-state';
import type { InsertText } from './ribbon-graphic-insert';
import { createInsertDialogFrame } from './insert-dialog-frame';

export function pickRibbonZoom(
	host: HTMLElement,
	target: RibbonGraphicTarget,
	kind: RibbonZoom,
	text: InsertText,
	current: () => RibbonGraphicTarget | undefined,
) {
	const doc = host.ownerDocument;
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
	const keepUnused = doc.createElement('input');
	keepUnused.type = 'checkbox';
	keepUnused.checked = true;
	if (multiple && target.sections.length) {
		const keepLabel = doc.createElement('label');
		keepLabel.append(
			keepUnused,
			doc.createTextNode(' ' + text('pptx.graphic.zoomKeepUnusedSections', 'Keep unused sections')),
		);
		body.append(keepLabel);
	}
	const sectionStarts = new Set<string>();
	const entries = options.map((choice) => {
		const label = doc.createElement('label');
		label.style.display = 'block';
		const input = doc.createElement('input');
		input.type = multiple ? 'checkbox' : 'radio';
		input.name = 'zoom-target';
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
		label.append(input, doc.createTextNode(` ${name}`));
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
				if (committing || !selected.length) return;
				committing = true;
				list.disabled = true;
				sync();
				void Promise.all(
					selected.map(async ({ choice, name }) => {
						const { width, height } = target.canvasSize;
						const ratio = Math.min(1, 360 / Math.max(width, height));
						const svg = SvgExporter.exportSlide(choice.slide, width, height);
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
								if (zoom) {
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
									live.insertSummary?.(prepared.result);
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
							resolve(createRibbonZoom(kind, resolved, live.canvasSize, heading));
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
				!committing && entries.some(({ input }) => input.checked),
			);
		list.addEventListener('change', sync);
		sync();
		frame.dialog.show();
	});
}
