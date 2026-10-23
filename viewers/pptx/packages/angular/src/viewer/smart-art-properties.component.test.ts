import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { PptxSmartArtData } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { SmartArtPropertiesComponent } from './smart-art-properties.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

const smartArtData: PptxSmartArtData = {
	resolvedLayoutType: 'list',
	nodes: [
		{ id: 'a', text: 'Alpha' },
		{ id: 'b', text: 'Bravo' },
	],
};

function mount(inputs: Record<string, unknown>) {
	TestBed.configureTestingModule({
		imports: [SmartArtPropertiesComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(SmartArtPropertiesComponent, { add: { inputs: Object.keys(inputs) } });
	const fixture = TestBed.createComponent(SmartArtPropertiesComponent);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const emitted: PptxSmartArtData[] = [];
	fixture.componentInstance.smartArtDataChange.subscribe((next) => emitted.push(next));
	fixture.detectChanges();
	const pane = (fixture.nativeElement as HTMLElement).querySelector('pptx-ui-smartart-text-pane')!;
	const inputsOf = () => [
		...pane.shadowRoot!.querySelectorAll<HTMLInputElement>('[data-testid="smartart-node-text"]'),
	];
	return { pane, inputs: inputsOf, emitted };
}

describe('smartArtPropertiesComponent (angular)', () => {
	it('feeds the shared text pane and emits the edits it makes', () => {
		const { inputs, emitted } = mount({ smartArtData, canEdit: true });
		expect(inputs().map((input) => input.value)).toEqual(['Alpha', 'Bravo']);
		inputs()[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
		expect(emitted).toHaveLength(1);
		expect(emitted[0]!.nodes.find((node) => node.id === 'b')?.parentId).toBe('a');
	});

	it('is read-only without edit rights', () => {
		const { inputs, emitted } = mount({ smartArtData, canEdit: false });
		expect(inputs().every((input) => input.disabled)).toBeTruthy();
		inputs()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(emitted).toEqual([]);
	});
});
