/**
 * Pointing slide masters at another theme part already in the package (Presentation properties >
 * Theme). `setPresentationTheme` used to rewrite only the masters' theme relationships: the
 * runtime's theme state still described the old part, so a scheme edit made after the switch (and
 * the save after it) went to the part the masters no longer use.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { buildTwoThemesDeck, TWO_THEMES } from '../../../__tests__/helpers/two-themes-deck';
import type { PptxData, PptxThemeColorScheme } from '../../types';
import { PptxHandler } from '../../PptxHandler';

const { theme1, theme2, master1, master2 } = TWO_THEMES;

async function open(bytes?: Uint8Array): Promise<{ handler: PptxHandler; data: PptxData }> {
	const source = bytes ?? (await buildTwoThemesDeck());
	const handler = new PptxHandler();
	const buffer = source.buffer.slice(
		source.byteOffset,
		source.byteOffset + source.byteLength,
	) as ArrayBuffer;
	return { handler, data: await handler.load(buffer) };
}

/** A theme part's name, accent 1 and heading face, as the saved package carries them. */
async function savedParts(handler: PptxHandler, data: PptxData) {
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const read = async (path: string) => {
		const xml = (await zip.file(path)?.async('string')) ?? '';
		return {
			name: /<a:theme\b[^>]*\bname="([^"]*)"/u.exec(xml)?.[1],
			accent1: /<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u.exec(xml)?.[1]?.toUpperCase(),
			major: /<a:majorFont>\s*<a:latin typeface="([^"]*)"/u.exec(xml)?.[1],
		};
	};
	return { theme1: await read(theme1.path), theme2: await read(theme2.path) };
}

const original = (theme: typeof theme1) => ({
	name: theme.name,
	accent1: theme.colors.accent1,
	major: theme.major,
});

function scheme(theme: typeof theme1, accent1: string): PptxThemeColorScheme {
	const colors = Object.fromEntries(
		Object.entries(theme.colors).map(([key, value]) => [key, `#${value}`]),
	);
	return { ...colors, accent1 } as unknown as PptxThemeColorScheme;
}

describe('getThemePartSchemes', () => {
	it("reads one theme part's colour scheme, font scheme and name", async () => {
		const { handler } = await open();
		const read = await handler.getThemePartSchemes(theme2.path);
		expect(read?.path).toBe(theme2.path);
		expect(read?.name).toBe('Ember');
		expect(read?.colorScheme.accent1?.replace('#', '').toUpperCase()).toBe('C0392B');
		expect(read?.colorScheme.dk2?.replace('#', '').toUpperCase()).toBe('4A1C14');
		expect(read?.fontScheme.majorFont?.latin).toBe('Courier New');
		expect(read?.fontScheme.minorFont?.latin).toBe('Arial');
		await expect(handler.getThemePartSchemes('ppt/theme/theme9.xml')).resolves.toBeUndefined();
	});

	it("carries the part's own format scheme", async () => {
		const { handler } = await open(await buildTwoThemesDeck({ formatSchemes: true }));
		const first = await handler.getThemePartSchemes(theme1.path);
		const second = await handler.getThemePartSchemes(theme2.path);
		expect(first?.formatScheme?.name).toBe('Tide');
		expect(second?.formatScheme?.name).toBe('Ember');
		expect(first?.formatScheme?.effectStyles[0]).not.toStrictEqual(
			second?.formatScheme?.effectStyles[0],
		);
	});

	it('changes nothing: an edit after a read still goes to the parts the masters use', async () => {
		const { handler, data } = await open();
		await handler.getThemePartSchemes(theme2.path);
		await handler.updateThemeFontScheme({ majorFont: { latin: 'Verdana' } });
		const parts = await savedParts(handler, data);
		expect(parts.theme1.major).toBe('Verdana');
		expect(parts.theme2.major).toBe('Verdana');
	});
});

describe('setPresentationTheme keeps the theme state on the part the masters use', () => {
	it('a colour scheme edit after switching every master goes to the new part only', async () => {
		const { handler, data } = await open();
		await handler.setPresentationTheme(theme2.path, true);
		await handler.updateThemeColorScheme(scheme(theme2, '#123456'));
		const parts = await savedParts(handler, data);
		expect(parts.theme2).toStrictEqual({ ...original(theme2), accent1: '123456' });
		expect(parts.theme1).toStrictEqual(original(theme1));
	});

	it('the first master alone: a font edit goes to the part both masters now use', async () => {
		const { handler, data } = await open();
		await handler.setPresentationTheme(theme2.path, false);
		await handler.updateThemeFontScheme({ majorFont: { latin: 'Verdana' } });
		const parts = await savedParts(handler, data);
		expect(parts.theme2.major).toBe('Verdana');
		expect(parts.theme1).toStrictEqual(original(theme1));
	});

	it('saves the masters on the new part, and a reload draws their slides with it', async () => {
		const { handler, data } = await open();
		await handler.setPresentationTheme(theme2.path, true);
		const reopened = await open(await handler.save(data.slides));
		expect(reopened.data.slideMasters?.map((master) => master.themePath)).toStrictEqual([
			theme2.path,
			theme2.path,
		]);
		const shape = reopened.data.slides[1]?.elements.find(
			(element) => element.name === TWO_THEMES.shapeName,
		);
		expect(shape?.shapeStyle?.fillColor?.replace('#', '').toUpperCase()).toBe('C0392B');
	});

	it('leaves the package alone when the theme part does not exist', async () => {
		const { handler, data } = await open();
		await handler.setPresentationTheme('ppt/theme/theme9.xml', true);
		const reopened = await open(await handler.save(data.slides));
		expect(reopened.data.slideMasters?.map((master) => master.themePath)).toStrictEqual([
			theme1.path,
			theme2.path,
		]);
	});
});

describe('setSlideMasterTheme', () => {
	it('re-points one master, and edits follow it', async () => {
		const { handler, data } = await open();
		await handler.setSlideMasterTheme(master2, theme1.path);
		await handler.updateThemeColorScheme(scheme(theme1, '#654321'));
		const parts = await savedParts(handler, data);
		expect(parts.theme1).toStrictEqual({ ...original(theme1), accent1: '654321' });
		expect(parts.theme2).toStrictEqual(original(theme2));
		const reopened = await open(await handler.save(data.slides));
		expect(reopened.data.slideMasters?.map((master) => master.themePath)).toStrictEqual([
			theme1.path,
			theme1.path,
		]);
	});

	it('puts a switched master back on its own part', async () => {
		const { handler, data } = await open();
		await handler.setPresentationTheme(theme2.path, true);
		await handler.setSlideMasterTheme(master1, theme1.path);
		await handler.updateThemeFontScheme({ majorFont: { latin: 'Verdana' } });
		const parts = await savedParts(handler, data);
		expect(parts.theme1.major).toBe('Verdana');
		expect(parts.theme2.major).toBe('Verdana');
		const reopened = await open(await handler.save(data.slides));
		expect(reopened.data.slideMasters?.map((master) => master.themePath)).toStrictEqual([
			theme1.path,
			theme2.path,
		]);
	});
});

describe('scheme edits that name their theme parts', () => {
	it('write only those parts, through the handler', async () => {
		const { handler, data } = await open();
		await handler.updateThemeColorScheme(scheme(theme1, '#654321'), [theme1.path]);
		await handler.updateThemeFontScheme({ majorFont: { latin: 'Verdana' } }, [theme1.path]);
		await handler.updateThemeName('Renamed', [theme1.path]);
		const parts = await savedParts(handler, data);
		expect(parts.theme1).toStrictEqual({ name: 'Renamed', accent1: '654321', major: 'Verdana' });
		expect(parts.theme2).toStrictEqual(original(theme2));
		const reopened = await open(await handler.save(data.slides));
		const heading = (index: number) => {
			const element = reopened.data.slides[index]?.elements.find((e) => e.name === 'Heading');
			return element && 'textSegments' in element
				? element.textSegments?.[0]?.style?.fontFamily
				: undefined;
		};
		expect([heading(0), heading(2)]).toStrictEqual(['Verdana', theme2.major]);
	});
});
