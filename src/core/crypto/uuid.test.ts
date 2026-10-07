import { afterEach, describe, expect, it, vi } from 'vitest';
import { createOfficeGuid } from './uuid';
const GUID_RE = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;

describe('createOfficeGuid', () => {
	it('generates a braced, uppercased GUID', () => {
		const id = createOfficeGuid();
		expect(id).toMatch(GUID_RE);
	});

	it('generates different ids on successive calls', () => {
		const a = createOfficeGuid();
		const b = createOfficeGuid();
		expect(a).not.toBe(b);
	});

	describe('without crypto.randomUUID', () => {
		afterEach(() => {
			vi.unstubAllGlobals();
		});

		it('falls back to crypto.getRandomValues (not Math.random) and still produces a valid GUID', () => {
			// Regression test for the CodeQL js/insecure-randomness finding: the
			// fallback path must source its bytes from the Web Crypto CSPRNG
			// whenever it is available, even when `randomUUID` specifically is not.
			const getRandomValues = vi.fn((array: Uint8Array) => {
				for (let i = 0; i < array.length; i++) {
					array[i] = i % 256;
				}
				return array;
			});
			vi.stubGlobal('crypto', { getRandomValues });

			const id = createOfficeGuid();

			expect(getRandomValues).toHaveBeenCalledWith(expect.any(Uint8Array));
			expect(id).toMatch(GUID_RE);
			// The variant nibble (`{XXXXXXXX-XXXX-4XXX-VXXX-XXXXXXXXXXXX}`, index
			// 20 once wrapped in braces) must stay in the RFC 4122 `10xx` range,
			// i.e. one of 8/9/A/B.
			expect(id[20]).toMatch(/[89AB]/);
		});

		it('still produces a valid GUID with no crypto object at all', () => {
			vi.stubGlobal('crypto', undefined);
			const id = createOfficeGuid();
			expect(id).toMatch(GUID_RE);
		});
	});
});
