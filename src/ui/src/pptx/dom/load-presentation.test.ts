// @vitest-environment happy-dom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import JSZip from 'jszip';
import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createProgressiveSlidePreview } from '../loader/progressive-slide-preview';
import { buildFieldSubstitutionContext } from '../render/field-context';
import { substituteFieldText } from '../render/text-field-substitution';
import { createDefaultRegistry, renderSlideStage } from './index';
import { loadPresentation } from './load-presentation';

const FIXTURE = resolve(__dirname, '../../../../../e2e/pptx/fixtures/sample-deck.pptx');

function readFixture(): ArrayBuffer {
	const bytes = readFileSync(FIXTURE);
	const buffer = new ArrayBuffer(bytes.byteLength);
	new Uint8Array(buffer).set(bytes);
	return buffer;
}

describe('loadPresentation (real .pptx happy path)', () => {
	it('keeps previews isolated from the editable deck and releases the preview nodes', async () => {
		const root = document.createElement('div');
		const host = document.createElement('div');
		host.dataset.pptxLoadPreviewHost = '';
		root.append(host);
		document.body.append(root);
		const preview = createProgressiveSlidePreview({
			getRoot: () => root,
			translate: createTranslator(),
			isCurrent: () => true,
			initialSlide: 0,
		});
		try {
			const loaded = await loadPresentation(readFixture(), { onSlideLoaded: preview.show });
			try {
				expect(host.querySelector('[data-pptx-opening-preview]')?.hasAttribute('aria-hidden')).toBe(
					true,
				);
				expect(host.querySelector('[data-element-id]')).toBeNull();
				expect(host.querySelectorAll('[data-pptx-preview-thumbnail]')).toHaveLength(6);
				expect(loaded.slides).toHaveLength(7);
			} finally {
				loaded.handler.dispose();
			}
		} finally {
			preview.dispose();
			root.remove();
		}
		expect(host.childElementCount).toBe(0);
	});

	it('applies the embedded viewer archive budget to a real deck', async () => {
		await expect(loadPresentation(readFixture(), { maxUncompressedBytes: 1 })).rejects.toThrow();
	});
	it('reports the parsed slides before the load finishes, once', async () => {
		const onSlidesParsed = vi.fn();
		const loaded = await loadPresentation(readFixture(), { onSlidesParsed });
		try {
			expect(onSlidesParsed).toHaveBeenCalledTimes(1);
			const parsed = onSlidesParsed.mock.calls[0]?.[0] as {
				slides: PptxSlide[];
				embeddedFonts: unknown[];
			};
			expect(parsed.slides.map((slide) => slide.id)).toEqual(
				loaded.slides.map((slide) => slide.id),
			);
			expect(Array.isArray(parsed.embeddedFonts)).toBeTruthy();
		} finally {
			loaded.handler.dispose();
		}
	});
	it('parses a real deck and renders its first slide through the registry', async () => {
		const loaded = await loadPresentation(readFixture());
		try {
			expect(loaded.slides.length).toBeGreaterThan(0);
			expect(loaded.canvasSize.width).toBeGreaterThan(0);
			expect(loaded.canvasSize.height).toBeGreaterThan(0);
			expect(Array.isArray(loaded.embeddedFonts)).toBeTruthy();
			expect(loaded.signatures).toBeUndefined();

			const stage = renderSlideStage({
				document,
				slide: loaded.slides[0],
				canvasSize: loaded.canvasSize,
				mediaDataUrls: loaded.mediaDataUrls,
				registry: createDefaultRegistry(),
				t: createTranslator(),
				scale: 0.5,
			});
			expect(stage.querySelectorAll('[data-element-id]').length).toBeGreaterThan(0);
		} finally {
			loaded.handler.dispose();
		}
	});

	// `p:presentation/@firstSlideNum="45"` must reach the field context, or a
	// `slidenum` field on slide 1 shows its position (1) instead of 45.
	it('carries the start slide number into the field context', async () => {
		const zip = await JSZip.loadAsync(readFixture());
		const xml = await zip.file('ppt/presentation.xml')!.async('string');
		zip.file(
			'ppt/presentation.xml',
			xml.replace('<p:presentation ', '<p:presentation firstSlideNum="45" '),
		);
		const loaded = await loadPresentation(await zip.generateAsync({ type: 'arraybuffer' }));
		try {
			expect(loaded.firstSlideNumber).toBe(45);
			const ctx = buildFieldSubstitutionContext({
				slide: loaded.slides[0],
				firstSlideNumber: loaded.firstSlideNumber,
			});
			expect(substituteFieldText('1', 'slidenum', ctx)).toBe('45');
		} finally {
			loaded.handler.dispose();
		}
	});

	it('rejects on invalid bytes without leaking the handler', async () => {
		await expect(loadPresentation(new ArrayBuffer(8))).rejects.toThrow();
	});
});
