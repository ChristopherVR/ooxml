import fs from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { PptxSlide, TextSegment } from '../../core/types';
import { requireFixture } from '../require-fixture';

/**
 * Issue 248: Insert > Date & Time with an "Available format" of just the year writes
 * `<a:fld type="datetimeyyyy">`. The type has to reach the renderer intact (it carries the
 * format) and has to survive a save, or the viewer can only show the default date.
 */
const fixture = requireFixture(
	path.resolve(__dirname, '../fixtures/issue-248-dynamic-date-field.pptx'),
);

function toBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

function dateFieldSegments(slide: PptxSlide | undefined): TextSegment[] {
	const found: TextSegment[] = [];
	for (const element of slide?.elements ?? []) {
		if (!('textSegments' in element)) continue;
		for (const segment of element.textSegments ?? []) {
			if (segment.fieldType?.toLowerCase().startsWith('datetime')) found.push(segment);
		}
	}
	return found;
}

describe('issue 248: format-encoded datetime field', () => {
	it('keeps the datetimeyyyy type and its cached text on load', async () => {
		const data = await new PptxHandler().load(toBuffer(fs.readFileSync(fixture)));
		const fields = dateFieldSegments(data.slides[0]);
		expect(fields.map((f) => f.fieldType)).toEqual(['datetimeyyyy']);
		expect(fields[0]?.text).toBe('2026');
	});

	it('writes the same field type back on save', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(toBuffer(fs.readFileSync(fixture)));
		const saved = await handler.save(data.slides);
		const reloaded = await new PptxHandler().load(toBuffer(saved));
		expect(dateFieldSegments(reloaded.slides[0]).map((f) => f.fieldType)).toEqual(['datetimeyyyy']);
	});
});
