import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import { setAfterAnimation } from '../../core/services/editor-animation-after-effect';
import { indexEffectNodes } from '../../core/services/animation-timing-tree';
import { applySoundToEffectCTn } from '../../core/services/animation-write-node-sound';
import type { XmlObject } from '../../core/types';

const fixture = readFileSync(
	new URL('../fixtures/e2e/animation-after-effect.pptx', import.meta.url),
);
const timing = (xml: string) => /<p:timing[\s\S]*?<\/p:timing>/u.exec(xml)?.[0];

it('preserves sound bytes, relationships and mixed sub-timing during clear and retiming', async () => {
	const zip = await JSZip.loadAsync(fixture);
	const sourceXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const tree = new XMLParser({ ignoreAttributes: false }).parse(sourceXml);
	const raw = tree['p:sld']['p:timing'] as XmlObject;
	const effect = indexEffectNodes(raw).find((ref) => ref.spid === '2')!;
	applySoundToEffectCTn(effect.cTn, { soundRId: 'rId100', soundName: 'fixture silence' });
	const sub = effect.cTn['p:subTnLst'] as XmlObject;
	sub['p:cmd'] = {
		'@_type': 'evt',
		'@_cmd': 'unmodeled',
		'p:cBhvr': { 'p:cTn': { '@_id': '999', '@_dur': '1' } },
	};
	zip.file('ppt/slides/slide1.xml', new XMLBuilder({ ignoreAttributes: false }).build(tree));
	const relPath = 'ppt/slides/_rels/slide1.xml.rels';
	const rels = (await zip.file(relPath)!.async('string')).replace(
		'</Relationships>',
		'<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/after-clear.wav"/></Relationships>',
	);
	zip.file(relPath, rels);
	const wave = Buffer.alloc(44 + 16000);
	wave.write('RIFF');
	wave.writeUInt32LE(wave.length - 8, 4);
	wave.write('WAVEfmt ', 8);
	wave.writeUInt32LE(16, 16);
	wave.writeUInt16LE(1, 20);
	wave.writeUInt16LE(1, 22);
	wave.writeUInt32LE(8000, 24);
	wave.writeUInt32LE(16000, 28);
	wave.writeUInt16LE(2, 32);
	wave.writeUInt16LE(16, 34);
	wave.write('data', 36);
	wave.writeUInt32LE(16000, 40);
	zip.file('ppt/media/after-clear.wav', wave);
	zip.file(
		'[Content_Types].xml',
		(await zip.file('[Content_Types].xml')!.async('string')).replace(
			'</Types>',
			'<Default Extension="wav" ContentType="audio/wav"/></Types>',
		),
	);
	const bytes = await zip.generateAsync({ type: 'nodebuffer' });
	for (const mode of ['noop', 'retime', 'clear']) {
		const handler = new PptxHandler();
		try {
			const data = await handler.load(
				bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
			);
			const slide = data.slides[0]!;
			const target = slide.nativeAnimations!.find((a) => a.afterAnimationAction === 'dimToColor')!;
			if (mode === 'clear')
				slide.animations = setAfterAnimation(slide.animations ?? [], target.targetId!, 'none');
			else if (mode === 'retime')
				slide.animations = slide.animations!.map((a) =>
					a.elementId === target.targetId ? { ...a, entrance: 'fadeIn', durationMs: 1234 } : a,
				);
			const saved = await handler.save(data.slides);
			const savedZip = await JSZip.loadAsync(saved);
			expect(await savedZip.file('ppt/media/after-clear.wav')!.async('nodebuffer')).toEqual(wave);
			expect(await savedZip.file(relPath)!.async('string')).toContain('Id="rId100"');
			expect(await savedZip.file(relPath)!.async('string')).toContain(
				'Target="../media/after-clear.wav"',
			);
			const savedXml = await savedZip.file('ppt/slides/slide1.xml')!.async('string');
			expect(savedXml).toContain('cmd="unmodeled"');
			if (mode === 'noop')
				expect(timing(savedXml)).toBe(
					timing(await zip.file('ppt/slides/slide1.xml')!.async('string')),
				);
			expect(savedXml).toContain('r:embed="rId100"');
			const reopened = await handler.load(
				saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
			);
			expect(
				reopened.slides[0]!.nativeAnimations!.find((a) => a.targetId === target.targetId)
					?.afterAnimationAction,
			).toBe(mode === 'clear' ? undefined : 'dimToColor');
		} finally {
			handler.dispose();
		}
	}
});

it('does not serialize imported control fallback or alter untouched native timing', async () => {
	const handler = new PptxHandler();
	try {
		const source = await handler.load(
			fixture.buffer.slice(fixture.byteOffset, fixture.byteOffset + fixture.byteLength),
		);
		expect(source.slides[0]!.animations?.every((a) => a.afterAnimation === undefined)).toBeTruthy();
		expect(
			source.slides[0]!.animations?.some((a) => a.nativeAfterAnimation === 'dimToColor'),
		).toBeTruthy();
		const saved = await handler.save(source.slides);
		const originalZip = await JSZip.loadAsync(fixture);
		const savedZip = await JSZip.loadAsync(saved);
		const before = await originalZip.file('ppt/slides/slide1.xml')!.async('string');
		const after = await savedZip.file('ppt/slides/slide1.xml')!.async('string');
		expect(timing(after)).toBe(timing(before));
		expect(after).not.toContain('pptx:animation');
	} finally {
		handler.dispose();
	}
});

it.each(['dimToColor', 'hideAfterAnimation', 'hideOnNextClick'] as const)(
	'clears imported %s with no preset rewrite and persists explicit none',
	async (action) => {
		const handler = new PptxHandler();
		try {
			const source = await handler.load(
				fixture.buffer.slice(fixture.byteOffset, fixture.byteOffset + fixture.byteLength),
			);
			const slide = source.slides[0]!;
			const target = slide.nativeAnimations!.find((a) => a.afterAnimationAction === action)!;
			const originalChildren = structuredClone(slide.rawTiming);
			slide.animations = setAfterAnimation(slide.animations ?? [], target.targetId!, 'none');
			const saved = await handler.save(source.slides);
			const reopened = await handler.load(
				saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
			);
			const parsed = reopened.slides[0]!;
			expect(
				parsed.nativeAnimations!.find((a) => a.targetId === target.targetId)?.afterAnimationAction,
			).toBeUndefined();
			expect(parsed.animations!.find((a) => a.elementId === target.targetId)?.afterAnimation).toBe(
				'none',
			);
			expect(
				parsed.animations!.find((a) => a.elementId === target.targetId)?.nativeAfterAnimation,
			).toBe('none');
			expect(parsed.nativeAnimations!.length).toBe(slide.nativeAnimations!.length);
			expect(slide.rawTiming).toEqual(originalChildren);
			expect(
				parsed
					.nativeAnimations!.filter((a) => a.targetId !== target.targetId)
					.map((a) => a.afterAnimationAction),
			).toEqual(
				slide
					.nativeAnimations!.filter((a) => a.targetId !== target.targetId)
					.map((a) => a.afterAnimationAction),
			);
		} finally {
			handler.dispose();
		}
	},
);

it('retimes an imported effect without deleting an untouched after-state', async () => {
	const handler = new PptxHandler();
	try {
		const source = await handler.load(
			fixture.buffer.slice(fixture.byteOffset, fixture.byteOffset + fixture.byteLength),
		);
		const slide = source.slides[0]!;
		const target = slide.nativeAnimations!.find((a) => a.afterAnimationAction === 'dimToColor')!;
		slide.animations = slide.animations!.map((a) =>
			a.elementId === target.targetId ? { ...a, entrance: 'fadeIn', durationMs: 1234 } : a,
		);
		const saved = await handler.save(source.slides);
		const reopened = await handler.load(
			saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength),
		);
		const native = reopened.slides[0]!.nativeAnimations!.find(
			(a) => a.targetId === target.targetId,
		)!;
		expect(native.afterAnimationAction).toBe('dimToColor');
		expect(native.afterAnimationColor).toBe(target.afterAnimationColor);
		expect(native.durationMs).toBe(1234);
	} finally {
		handler.dispose();
	}
});
