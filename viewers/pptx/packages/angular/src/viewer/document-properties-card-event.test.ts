import '@angular/compiler';
import { Injector, runInInjectionContext, signal } from '@angular/core';
import { describe, expect, it, vi } from 'vitest';

import { DocumentPropertiesCardComponent } from './document-properties-card.component';
import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { ViewerDocumentPropertiesService } from './viewer-document-properties.service';

describe('propertiesChange from the inspector card', () => {
	it('reports a core-properties edit made in the inspector', () => {
		const coreProperties = signal<Record<string, string> | undefined>(undefined);
		const emitPropertiesChange = vi.fn();
		const docProperties = { notifyCoreChange: emitPropertiesChange };
		const injector = Injector.create({
			providers: [
				{
					provide: LoadContentService,
					useValue: {
						coreProperties,
						appProperties: signal(undefined),
						customProperties: signal([]),
					},
				},
				{ provide: EditorStateService, useValue: { dirty: signal(false) } },
				{ provide: ViewerDocumentPropertiesService, useValue: docProperties },
			],
		});
		const card = runInInjectionContext(injector, () => new DocumentPropertiesCardComponent());

		(
			card as unknown as {
				onCoreChange(event: Event, key: 'title' | 'creator'): void;
			}
		).onCoreChange({ target: { value: 'Quarterly review' } } as unknown as Event, 'title');

		expect(coreProperties()).toStrictEqual({ title: 'Quarterly review' });
		expect(emitPropertiesChange).toHaveBeenCalledExactlyOnceWith({ title: 'Quarterly review' });
	});
});
