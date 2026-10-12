import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { adaptDocumentModel } from './adapter';

it.each([undefined, 11, 13, 16])(
	'preserves unverified compatibility mode %s without assuming modern note layout',
	async (mode) => {
		const { model } = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL('../__fixtures__/footnote-columns/two-two-twenty.docx', import.meta.url),
				),
			),
		);
		if (mode === undefined) delete model.compatibilityMode;
		else model.compatibilityMode = mode;
		const warnings: string[] = [];
		const layout = adaptDocumentModel(model, (message) => warnings.push(message));
		expect(model.sections![0]!.footnoteColumns).toBe(2);
		expect(layout.sections[0]!.footnoteColumns).toBeUndefined();
		expect(layout.sections[0]!.modernFootnoteSeparator).toBeUndefined();
		expect(warnings.some((message) => message.includes('has not been verified'))).toBe(true);
	},
);
