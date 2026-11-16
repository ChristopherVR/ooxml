/**
 * Source-text guards for the Angular Fill section (no TestBed here, see
 * `vitest.config.ts`): the fields are the shared fill descriptors, and the
 * stop rows, swatches and picture source report the shared fill actions.
 */
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';

const section = componentSource(import.meta.dirname, 'shape-fill-section.component.ts');
const gradient = componentSource(import.meta.dirname, 'gradient-picker.component.ts');
const pattern = componentSource(import.meta.dirname, 'pattern-fill-panel.component.ts');

describe('shapeFillSectionComponent', () => {
	it('draws the shared fill descriptors around the projected solid colour', () => {
		expect(section).toContain('SHAPE_FILL_MODE_SECTIONS');
		expect(section).toContain('SHAPE_FILL_OPTIONS_SECTIONS');
		expect(section.indexOf('[sections]="modeSections"')).toBeLessThan(
			section.indexOf('<ng-content />'),
		);
		expect(section.indexOf('<ng-content />')).toBeLessThan(
			section.indexOf('[sections]="optionSections"'),
		);
		expect(section).toContain('@if (view().stops)');
		expect(section).toContain('@if (view().pattern)');
		expect(section).toContain("shapeFillPatch(this.style(), { type: 'imageUrl', url })");
	});

	it('draws the stop rows and swatches from the shared view and reports shared actions', () => {
		expect(gradient).toContain('resolveShapeFillView');
		expect(gradient).toContain('[selectedRef]="stop.colorRef"');
		expect(gradient).toContain('(pick)="onStopThemePick($event, stop.index)"');
		expect(gradient).toContain(
			"this.act({ type: 'stopColor', index, hex: commit.hex, ref: commit.ref });",
		);
		// A native pick carries no theme ref, so the shared action clears the stop's.
		expect(gradient).toContain("this.act({ type: 'stopColor', index, hex: target.value });");
		expect(pattern).toContain('data-testid="fill-pattern-swatch"');
		expect(pattern).toContain("this.act({ type: 'patternPreset', preset });");
	});
});
