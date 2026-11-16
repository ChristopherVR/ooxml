/**
 * accessibility-text-panel.component.test.ts
 *
 * `AccessibilityTextPanelComponent` is the alt text / title editor for a
 * plain shape, text box, connector, or any graphic-frame kind
 * (table/chart/smartArt/media/ole), at parity with React's
 * `AccessibilityTextSection` and Vue's `AccessibilityPanel.vue`.
 *
 * No Angular TestBed here (see `media-properties-panel.component.test.ts`):
 * `.supports()` is a plain static method testable without DI, and the
 * remaining assertions pin the source to the wiring the template/computed
 * rely on, importing the renderer through the public `ooxml-ui/pptx` dependency.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { AccessibilityTextPanelComponent } from './accessibility-text-panel.component';

function el(type: PptxElement['type']): PptxElement {
	return { id: 'el-1', type, x: 0, y: 0, width: 10, height: 10 } as unknown as PptxElement;
}

describe('accessibilityTextPanelComponent.supports', () => {
	it('supports a plain shape, text box and connector', () => {
		expect(AccessibilityTextPanelComponent.supports(el('shape'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('text'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('connector'))).toBeTruthy();
	});

	it('excludes a picture, whose own alt text field lives in ImagePropertiesPanelComponent', () => {
		expect(AccessibilityTextPanelComponent.supports(el('image'))).toBeFalsy();
		expect(AccessibilityTextPanelComponent.supports(el('picture'))).toBeFalsy();
	});

	it('supports every graphic-frame kind (table/chart/smartArt/media/ole)', () => {
		expect(AccessibilityTextPanelComponent.supports(el('table'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('chart'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('smartArt'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('ole'))).toBeTruthy();
		expect(AccessibilityTextPanelComponent.supports(el('media'))).toBeTruthy();
	});

	it('excludes a kind with neither field, like a group', () => {
		expect(AccessibilityTextPanelComponent.supports(el('group'))).toBeFalsy();
	});
});

describe('accessibilityTextPanelComponent source wiring', () => {
	const source = readFileSync(
		path.join(__dirname, 'accessibility-text-panel.component.ts'),
		'utf8',
	);

	it('reads its fields through the public UI descriptor', () => {
		expect(source).toMatch(/getNonVisualDescriptionFields/);
		expect(source).toContain("from 'ooxml-ui/pptx'");
		expect(source).not.toContain("from '../internal/shared'");
	});

	it('draws the shared alt text descriptors with the generic renderer', () => {
		expect(source).toContain('ELEMENT_ALT_TEXT_SECTIONS');
		expect(source).toContain('<pptx-inspector-fields');
		expect(source).toContain('[disabled]="!canEdit()"');
	});
});

describe('picture panels draw the shared picture descriptors', () => {
	const image = readFileSync(path.join(__dirname, 'image-properties-panel.component.ts'), 'utf8');
	const crop = readFileSync(path.join(__dirname, 'image-crop-wash-panel.component.ts'), 'utf8');

	it('draws alt text and the adjustments from the shared sections', () => {
		expect(image).toContain('ELEMENT_ALT_TEXT_SECTIONS');
		expect(image).toContain('PICTURE_ADJUSTMENT_SECTIONS');
		// Transparency is the shared, inverted reading of alphaModFix, not the raw value.
		expect(image).not.toContain("key: 'alphaModFix'");
	});

	it('draws the crop insets from the shared section, gated on croppable', () => {
		expect(crop).toContain('PICTURE_CROP_SECTIONS');
		expect(crop).toContain('[disabled]="!croppable()"');
		expect(crop).not.toContain('Crop {{ side }}');
	});
});

describe('inspectorPanelComponent wiring for accessibility text', () => {
	const source = readFileSync(path.join(__dirname, 'inspector-panel.component.ts'), 'utf8');

	it('mounts pptx-accessibility-text-panel gated on accessibilityTextEl', () => {
		expect(source).toContain('pptx-accessibility-text-panel');
		expect(source).toMatch(/\[element\]="a11yEl"\s*\[canEdit\]="canEdit\(\)"/);
		expect(source).toMatch(/accessibilityTextEl\(\);\s*as a11yEl/);
		expect(source).toMatch(/AccessibilityTextPanelComponent\.supports\(this\.el\(\)\)/);
	});
});
