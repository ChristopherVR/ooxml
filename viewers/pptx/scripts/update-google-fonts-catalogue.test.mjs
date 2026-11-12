import assert from 'node:assert/strict';
import { test } from 'node:test';

import { familyVariantMask, generateCatalogue } from './update-google-fonts-catalogue.mjs';

const faces = (mask) => {
	const result = [];
	for (const italic of [false, true]) {
		for (let weight = 100; weight <= 900; weight += 100) {
			if (mask & (1 << (weight / 100 - 1 + (italic ? 9 : 0)))) {
				result.push({ weight, italic });
			}
		}
	}
	return result;
};
const entry = (fonts, axes = []) => ({ family: 'Example', fonts, axes });
const normal = (...weights) => weights.map((weight) => ({ weight, italic: false }));

test('expands variable weight ranges for each existing style, capped at standard weights', () => {
	assert.deepEqual(
		faces(
			familyVariantMask(
				entry({ 100: {}, '100i': {}, 1000: {} }, [{ tag: 'wght', min: 100, max: 1000 }]),
			),
		),
		[
			...normal(100, 200, 300, 400, 500, 600, 700, 800, 900),
			...normal(100, 200, 300, 400, 500, 600, 700, 800, 900).map((face) => ({
				...face,
				italic: true,
			})),
		],
	);
});

test('respects fractional range bounds and does not invent an italic style', () => {
	assert.deepEqual(
		faces(familyVariantMask(entry({ 250: {} }, [{ tag: 'wght', min: 250, max: 850 }]))),
		normal(300, 400, 500, 600, 700, 800),
	);
});

test('preserves sparse static weights and differing italic availability', () => {
	assert.deepEqual(faces(familyVariantMask(entry({ 200: {}, 400: {}, 700: {}, '400i': {} }))), [
		...normal(200, 400, 700),
		{ weight: 400, italic: true },
	]);
});

test('keeps singleton weights without inventing a regular face', () => {
	assert.deepEqual(faces(familyVariantMask(entry({ 700: {} }))), normal(700));
	assert.deepEqual(faces(familyVariantMask(entry({ '400i': {} }))), [
		{ weight: 400, italic: true },
	]);
});

test('ignores nonstandard weights and unrelated variable axes', () => {
	assert.deepEqual(
		faces(
			familyVariantMask(
				entry({ 350: {}, 400: {}, 950: {}, '1000i': {} }, [{ tag: 'wdth', min: 50, max: 200 }]),
			),
		),
		normal(400),
	);
});

test('rejects malformed faces and weight axes instead of silently emitting inaccurate data', () => {
	assert.throws(() => familyVariantMask({ family: 'Missing' }), /missing font faces/);
	assert.throws(() => familyVariantMask(entry({ regular: {} })), /unrecognised face/);
	assert.throws(() => familyVariantMask(entry({ 400: {} }, [{ tag: 'wght', min: 900, max: 100 }])));
	assert.throws(() =>
		familyVariantMask(entry({ 400: {} }, [{ tag: 'wght', min: '100', max: 900 }])),
	);
});

test('emits deterministic sorted names and aligned compact masks, merging duplicate families', () => {
	const metadata = {
		familyMetadataList: [
			{ family: 'Zulu', fonts: { 700: {} } },
			{ family: '  Alpha  ', fonts: { 400: {} } },
			{ family: 'Alpha', fonts: { '400i': {} } },
			{ family: ' ', fonts: {} },
		],
	};
	const source = generateCatalogue(metadata, '2026-11-12', 2);
	assert.match(source, /\t'Alpha',\n\t'Zulu',/u);
	assert.match(source, /GOOGLE_FONTS_CATALOGUE_DATE = '2026-11-12'/u);
	const packed = /GOOGLE_FONTS_VARIANT_MASKS =\n\t'([^']*)';/u.exec(source)?.[1];
	assert.equal(packed?.length, 8);
	assert.equal(parseInt(packed.slice(0, 4), 36), familyVariantMask(entry({ 400: {}, '400i': {} })));
	assert.equal(parseInt(packed.slice(4, 8), 36), familyVariantMask(entry({ 700: {} })));
	assert.equal(
		source,
		generateCatalogue(
			{ familyMetadataList: [...metadata.familyMetadataList].reverse() },
			'2026-11-12',
			2,
		),
	);
});

test('escapes family names in generated TypeScript and refuses suspicious catalogue shrinkage', () => {
	assert.match(
		generateCatalogue(
			{ familyMetadataList: [{ family: "A'B\\C", fonts: { 400: {} } }] },
			'2026-11-12',
			1,
		),
		/A\\'B\\\\C/u,
	);
	assert.throws(
		() => generateCatalogue({ familyMetadataList: [entry({ 400: {} })] }, '2026-11-12'),
		/refusing to shrink/,
	);
});
