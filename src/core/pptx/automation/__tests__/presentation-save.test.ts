import { describe, expect, it } from 'vitest';
import { pptxTools } from '../catalogue/index';
import { createTestPptxBytes, loadTestPptxData } from './helpers/create-test-pptx';

async function edit(bytes: Uint8Array, name: string, params: Record<string, unknown>) {
	const tool = pptxTools.find((candidate) => candidate.name === name);
	if (tool?.kind !== 'edit') throw new Error(`Missing edit tool ${name}`);
	const result = await tool.run(bytes, params);
	if (!result.bytes) throw new Error(`${name} did not save`);
	return result.bytes;
}

describe('presentation automation save', () => {
	it('persists section creation, membership, ordering and removal through registered tools', async () => {
		let bytes = await createTestPptxBytes(3);
		bytes = await edit(bytes, 'manage_sections', {
			action: 'add',
			name: 'Opening',
			slideIndices: [0, 1],
		});
		let data = await loadTestPptxData(bytes);
		expect(data.sections).toHaveLength(1);
		expect(data.sections?.[0]).toMatchObject({
			name: 'Opening',
			slideIds: data.slides.slice(0, 2).map((s) => s.id),
		});
		const opening = data.sections![0].id;
		bytes = await edit(bytes, 'manage_sections', {
			action: 'add',
			name: 'Closing',
			slideIndices: [2],
		});
		data = await loadTestPptxData(bytes);
		const closing = data.sections![1].id;
		bytes = await edit(bytes, 'manage_sections', {
			action: 'moveSlides',
			sectionId: closing,
			slideIndices: [1],
		});
		data = await loadTestPptxData(bytes);
		expect(data.sections?.find((s) => s.id === closing)?.slideIds).toHaveLength(2);
		bytes = await edit(bytes, 'manage_sections', {
			action: 'reorder',
			sectionIds: [closing, opening],
		});
		data = await loadTestPptxData(bytes);
		expect(data.sections?.map((s) => s.name)).toEqual(['Closing', 'Opening']);
		for (const id of [closing, opening])
			bytes = await edit(bytes, 'manage_sections', { action: 'remove', sectionId: id });
		expect((await loadTestPptxData(bytes)).sections ?? []).toEqual([]);
	});

	it.each(['presented', 'browsed', 'kiosk'] as const)(
		'persists %s slideshow options and keeps false values in later updates',
		async (showType) => {
			let bytes = await createTestPptxBytes(3);
			const options = {
				showType,
				loopContinuously: true,
				showWithNarration: false,
				showWithAnimation: false,
				advanceMode: 'manual',
				showSlidesMode: 'range',
				showSlidesFrom: 2,
				showSlidesTo: 3,
				penColor: '#FF0000',
			};
			bytes = await edit(bytes, 'update_presentation_properties', options);
			expect((await loadTestPptxData(bytes)).presentationProperties).toMatchObject(options);
			bytes = await edit(bytes, 'update_presentation_properties', { loopContinuously: false });
			expect((await loadTestPptxData(bytes)).presentationProperties).toMatchObject({
				...options,
				loopContinuously: false,
			});
		},
	);

	it('selects a custom show and clears the selection when that show is deleted', async () => {
		let bytes = await createTestPptxBytes(3);
		await expect(
			edit(bytes, 'update_presentation_properties', {
				showSlidesMode: 'customShow',
				showSlidesCustomShowId: 'missing',
			}),
		).rejects.toThrow('existing custom show');
		bytes = await edit(bytes, 'manage_custom_shows', {
			action: 'create',
			name: 'Short',
			slideIndices: [0, 2],
		});
		const id = (await loadTestPptxData(bytes)).customShows![0].id;
		bytes = await edit(bytes, 'update_presentation_properties', {
			showSlidesMode: 'customShow',
			showSlidesCustomShowId: id,
		});
		expect((await loadTestPptxData(bytes)).presentationProperties).toMatchObject({
			showSlidesMode: 'customShow',
			showSlidesCustomShowId: id,
		});
		bytes = await edit(bytes, 'manage_custom_shows', { action: 'delete', showId: id });
		const data = await loadTestPptxData(bytes);
		expect(data.customShows ?? []).toEqual([]);
		expect(data.presentationProperties?.showSlidesMode).toBe('all');
		expect(data.presentationProperties?.showSlidesCustomShowId).toBeUndefined();
	});
});
