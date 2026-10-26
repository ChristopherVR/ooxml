import { PptxLoadLimitError, ZipBombError } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n/translator';
import { translationsEn } from '../i18n/translations-en';
import { formatLoadLimitBytes, LOAD_LIMIT_MESSAGE_KEYS, pptxLoadErrorMessage } from './load-limits';

const t = createTranslator('en');

describe('pptxLoadErrorMessage', () => {
	it('names the slide limit and the observed count', () => {
		expect(pptxLoadErrorMessage(new PptxLoadLimitError('raw', 'maxSlides', 2, 3), t)).toBe(
			'This presentation has 3 slides, more than the 2 this viewer accepts.',
		);
	});

	it('formats byte limits with a unit', () => {
		expect(
			pptxLoadErrorMessage(
				new PptxLoadLimitError('raw', 'maxInputBytes', 1024 * 1024, 3 * 1024 * 1024),
				t,
			),
		).toBe('This file is 3 MB, larger than the 1 MB this viewer accepts.');
	});

	it('reads the zip guard limits from a ZipBombError', () => {
		expect(
			pptxLoadErrorMessage(
				new ZipBombError('raw', { limit: 10, loadLimit: 'maxEntries', entryCount: 12, actual: 12 }),
				t,
			),
		).toBe('This presentation contains 12 parts, more than the 10 this viewer accepts.');
		expect(
			pptxLoadErrorMessage(
				new ZipBombError('raw', {
					limit: 2048,
					loadLimit: 'maxUncompressedBytes',
					uncompressedBytes: 4096,
					actual: 4096,
				}),
				t,
			),
		).toBe('This presentation unpacks to at least 4 KB, more than the 2 KB this viewer accepts.');
	});

	it('keeps the message of any other error, including a forged zip', () => {
		expect(pptxLoadErrorMessage(new Error('Invalid PPTX'), t)).toBe('Invalid PPTX');
		expect(pptxLoadErrorMessage(new ZipBombError('forged', { limit: 1 }), t)).toBe('forged');
		expect(pptxLoadErrorMessage('plain', t)).toBe('plain');
	});

	it('passes the localized key and parameters to the binding translator', () => {
		const calls: unknown[] = [];
		const message = pptxLoadErrorMessage(
			new PptxLoadLimitError('raw', 'maxShapes', 100, 101),
			(key, params) => {
				calls.push([key, params]);
				return 'localized';
			},
		);
		expect(message).toBe('localized');
		expect(calls).toStrictEqual([
			['pptx.viewer.loadLimit.maxShapes', { actual: '101', maximum: '100' }],
		]);
	});

	it('has an English message for every limit', () => {
		for (const key of Object.values(LOAD_LIMIT_MESSAGE_KEYS)) {
			expect(translationsEn[key]).toMatch(/\{\{actual\}\}.*\{\{maximum\}\}/u);
		}
	});
});

describe('formatLoadLimitBytes', () => {
	it('uses B, KB and MB', () => {
		expect(formatLoadLimitBytes(512)).toBe('512 B');
		expect(formatLoadLimitBytes(1536)).toBe('1.5 KB');
		expect(formatLoadLimitBytes(500 * 1024 * 1024)).toBe('500 MB');
	});
});
