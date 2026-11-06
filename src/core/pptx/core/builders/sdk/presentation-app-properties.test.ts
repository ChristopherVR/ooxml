/**
 * `PresentationOptions.appProperties`: application properties set when a
 * presentation is created, saved through the public SDK and reloaded.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxHandlerSaveOptions } from '../../core/types';
import { Presentation } from './Presentation';
import type { PresentationOptions } from './types';

/** Create, add one slide, save, and return `docProps/app.xml` and the reloaded properties. */
async function createAndSave(
	options: PresentationOptions,
	saveOptions?: PptxHandlerSaveOptions,
): Promise<{ appXml: string; app: Record<string, unknown> }> {
	const { handler, createSlide } = await PptxHandler.create(options);
	const bytes = await handler.save([createSlide('Blank').build()], saveOptions);
	const zip = await JSZip.loadAsync(bytes);
	const appXml = (await zip.file('docProps/app.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	return { appXml, app: { ...loaded.appProperties } };
}

describe('PresentationOptions.appProperties', () => {
	it('writes Application, AppVersion and Company given at create time', async () => {
		const { appXml, app } = await createAndSave({
			appProperties: { application: 'ReportOne', appVersion: '2.0100', company: 'Acme' },
		});
		expect(appXml).toContain('<Application>ReportOne</Application>');
		expect(appXml).toContain('<AppVersion>2.0100</AppVersion>');
		expect(appXml).toContain('<Company>Acme</Company>');
		expect(app).toMatchObject({ application: 'ReportOne', appVersion: '2.0100', company: 'Acme' });
	});

	it('keeps the documented defaults when the option is omitted', async () => {
		const { appXml, app } = await createAndSave({});
		expect(appXml).toContain('<Application>pptx-viewer-sdk</Application>');
		expect(appXml).toContain('<AppVersion>1.0000</AppVersion>');
		expect(appXml).not.toContain('<Company>');
		expect(app.company).toBeUndefined();
	});

	it('keeps the default for a field the option leaves out', async () => {
		const { appXml } = await createAndSave({ appProperties: { company: 'Acme' } });
		expect(appXml).toContain('<Application>pptx-viewer-sdk</Application>');
		expect(appXml).toContain('<AppVersion>1.0000</AppVersion>');
		expect(appXml).toContain('<Company>Acme</Company>');
	});

	it('writes the other text fields and escapes markup', async () => {
		const { appXml, app } = await createAndSave({
			appProperties: {
				presentationFormat: 'Widescreen',
				manager: 'A & B <Ltd>',
				template: 'Brand.potx',
			},
		});
		expect(appXml).toContain('<Manager>A &amp; B &lt;Ltd&gt;</Manager>');
		expect(app).toMatchObject({
			presentationFormat: 'Widescreen',
			manager: 'A & B <Ltd>',
			template: 'Brand.potx',
		});
	});

	it('leaves an element out for an empty string', async () => {
		const { appXml, app } = await createAndSave({ appProperties: { application: '' } });
		expect(appXml).not.toContain('<Application>');
		expect(app.application).toBeUndefined();
	});

	it('trims values the way the save option does', async () => {
		const { appXml } = await createAndSave({
			appProperties: { company: '  Acme  ', appVersion: ' 16.0000 ' },
		});
		expect(appXml).toContain('<Company>Acme</Company>');
		expect(appXml).toContain('<AppVersion>16.0000</AppVersion>');
	});

	it('lets a save option win field by field over the create-time value', async () => {
		const { appXml, app } = await createAndSave(
			{ appProperties: { application: 'ReportOne', company: 'Acme' } },
			{ appProperties: { company: 'Globex' } },
		);
		expect(appXml).toContain('<Application>ReportOne</Application>');
		expect(appXml).toContain('<Company>Globex</Company>');
		expect(app).toMatchObject({ application: 'ReportOne', company: 'Globex' });
	});

	it('still recomputes the slide count on save', async () => {
		const { appXml } = await createAndSave({ appProperties: { company: 'Acme' } });
		expect(appXml).toContain('<Slides>1</Slides>');
	});

	it('is exposed on the loaded data right after create', async () => {
		const { data } = await PptxHandler.create({ appProperties: { company: 'Acme' } });
		expect(data.appProperties?.company).toBe('Acme');
	});

	it('is accepted by Presentation.create too', async () => {
		const pptx = await Presentation.create({ appProperties: { application: 'ReportOne' } });
		pptx.addSlide('Blank');
		const zip = await JSZip.loadAsync(await pptx.save());
		const appXml = (await zip.file('docProps/app.xml')?.async('string')) ?? '';
		expect(appXml).toContain('<Application>ReportOne</Application>');
	});

	it.each(['16', '16.0', '16.00000', 'v16.0000', 'latest'])(
		'refuses the AppVersion %s, which PowerPoint reports as corrupt',
		async (appVersion) => {
			await expect(PptxHandler.create({ appProperties: { appVersion } })).rejects.toThrow(
				RangeError,
			);
		},
	);

	it('refuses a field that is not set at create time', async () => {
		const appProperties = { slides: 4 } as unknown as PresentationOptions['appProperties'];
		await expect(PptxHandler.create({ appProperties })).rejects.toThrow(
			/appProperties\.slides cannot be set/u,
		);
	});

	it('refuses a value that is not a string', async () => {
		const appProperties = { company: 42 } as unknown as PresentationOptions['appProperties'];
		await expect(PptxHandler.create({ appProperties })).rejects.toThrow(RangeError);
	});
});
