import { expect, it } from 'vitest';
import { composeAffine, IDENTITY_AFFINE } from './affine';

it('composes child translation through a rotated parent, preserving order', () => {
	const parent = [0, 1, -1, 0, 4, 5] as const;
	const child = [1, 0, 0, 1, 2, 3] as const;
	expect(composeAffine(parent, child)).toEqual([0, 1, -1, 0, 1, 7]);
	expect(composeAffine(child, parent)).toEqual([0, 1, -1, 0, 6, 8]);
	expect(composeAffine(IDENTITY_AFFINE, parent)).toEqual(parent);
});
