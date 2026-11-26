import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	loadDocx,
	resolveWebVideo,
	signatureLineImage,
	signatureLinePlaceholderPng,
	webVideoImage,
	webVideoPosterPng,
} from './index';
import { createDocument } from './model';
import { saveDocx } from './save';
import { expectParagraph, must } from './test-support/access';

const PNG_TYPE = 'image/png';

describe('signature line round trip', () => {
	it('writes a VML signature line Word recognises and reads it back', async () => {
		const model = createDocument();
		const image = signatureLineImage(
			{
				id: '{11111111-2222-3333-4444-555555555555}',
				suggestedSigner: 'Ada Lovelace',
				suggestedSignerTitle: 'Analyst',
				suggestedSignerEmail: 'ada@example.com',
				instructions: 'Sign here',
				allowComments: true,
				showSignDate: false,
			},
			{ partName: 'word/media/image1.png', contentType: PNG_TYPE },
		);
		expectParagraph(model.blocks[0]).runs = [{ text: '', image }];
		const bytes = await saveDocx(
			model,
			new Map([
				['word/media/image1.png', { bytes: signatureLinePlaceholderPng(), contentType: PNG_TYPE }],
			]),
		);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await must(zip.file('word/document.xml')).async('string');
		expect(xml).toContain('issignatureline="t"');
		expect(xml).toContain('o:suggestedsigner="Ada Lovelace"');
		expect(xml).toContain('o:suggestedsigner2="Analyst"');
		expect(xml).toContain('signinginstructionsset="t"');
		expect(xml).toContain('showsigndate="f"');
		expect(zip.file('word/media/image1.png')).toBeTruthy();
		const again = await loadDocx(bytes);
		const run = must(expectParagraph(again.model.blocks[0]).runs[0]);
		expect(run.image?.signatureLine).toMatchObject({
			id: '{11111111-2222-3333-4444-555555555555}',
			suggestedSigner: 'Ada Lovelace',
			suggestedSignerEmail: 'ada@example.com',
			instructions: 'Sign here',
			allowComments: true,
			showSignDate: false,
		});
		// An unchanged save keeps the shape as written.
		const resaved = await JSZip.loadAsync(await again.save(again.model));
		expect(await must(resaved.file('word/document.xml')).async('string')).toContain(
			'issignatureline="t"',
		);
	});
});

describe('online video', () => {
	it('accepts only allow-listed https embeds and makes no request', () => {
		const allow = ['video.example.com'];
		const code =
			'<iframe width="560" src="https://player.video.example.com/embed/1?a=1&b=2"></iframe>';
		expect(resolveWebVideo(code, allow)).toMatchObject({ ok: true });
		expect(resolveWebVideo('https://video.example.com/embed/1', allow)).toMatchObject({ ok: true });
		expect(resolveWebVideo('https://evil.example.org/embed/1', allow)).toEqual({
			ok: false,
			reason: 'host',
		});
		expect(resolveWebVideo('http://video.example.com/embed/1', allow)).toEqual({
			ok: false,
			reason: 'invalid',
		});
		expect(resolveWebVideo('not a url', allow)).toEqual({ ok: false, reason: 'invalid' });
		expect(resolveWebVideo('https://video.example.com/embed/1', [])).toMatchObject({
			reason: 'host',
		});
	});

	it('writes wp15:webVideoPr on the poster picture and reads it back', async () => {
		const result = resolveWebVideo('https://video.example.com/embed/1?a=1&b=2', [
			'video.example.com',
		]);
		if (!result.ok) throw new Error('video refused');
		const model = createDocument();
		expectParagraph(model.blocks[0]).runs = [
			{
				text: '',
				image: webVideoImage(result.spec, {
					partName: 'word/media/image2.png',
					contentType: PNG_TYPE,
				}),
			},
		];
		const bytes = await saveDocx(
			model,
			new Map([['word/media/image2.png', { bytes: webVideoPosterPng(), contentType: PNG_TYPE }]]),
		);
		const xml = await must((await JSZip.loadAsync(bytes)).file('word/document.xml')).async(
			'string',
		);
		expect(xml).toContain('wp15:webVideoPr');
		expect(xml).toContain('{C183D7F6-B498-43B3-948B-1728B52AA6E4}');
		expect(xml).toContain('Online Media');
		const again = await loadDocx(bytes);
		const run = must(expectParagraph(again.model.blocks[0]).runs[0]);
		expect(run.image?.webVideo?.url).toBe('https://video.example.com/embed/1?a=1&b=2');
		expect(run.image?.webVideo?.embeddedHtml).toContain('<iframe');
	});
});
