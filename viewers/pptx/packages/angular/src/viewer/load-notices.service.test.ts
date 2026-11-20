/**
 * load-notices.service.test.ts: the read-only recommendation half of
 * `LoadNoticesService` (in particular the modify-password unlock prompt) and
 * the signature load notice.
 *
 * Bare injector (no TestBed): `LoadContentService` is provided as a minimal
 * stub exposing just the `parsedData`, `slides` and `signatureSummary` signals
 * `LoadNoticesService` reads; without `TranslateService` the notice is English.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { PptxData, SignatureSummary } from 'pptx-viewer-core';
import { createModifyVerifier, createSaltlessModifyVerifierForTesting } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { LoadContentService } from './load-content.service';
import { LoadNoticesService } from './load-notices.service';

function createService(parsedData: PptxData | undefined): LoadNoticesService {
	const injector = Injector.create({
		providers: [
			{
				provide: LoadContentService,
				useValue: {
					parsedData: signal(parsedData),
					slides: signal([]),
					signatureSummary: signal(undefined),
				},
			},
		],
	});
	return runInInjectionContext(injector, () => new LoadNoticesService());
}

describe('loadNoticesService', () => {
	it('shows no banner and stays unlocked without a recommendation', () => {
		const service = createService(undefined);
		expect(service.bannerActive()).toBeFalsy();
		expect(service.lockActive()).toBeFalsy();
	});

	it('recommends read-only and locks editing for a modifyVerifier deck', () => {
		const service = createService({
			modifyVerifier: { hashData: 'abc', algorithmName: 'SHA-512' },
		} as unknown as PptxData);
		expect(service.recommendation().kind).toBe('modifyVerifier');
		expect(service.bannerActive()).toBeTruthy();
		expect(service.lockActive()).toBeTruthy();
	});

	describe('password-protected modifyVerifier', () => {
		it('editAnyway opens the password prompt instead of unlocking', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			expect(service.recommendation().requiresPassword).toBeTruthy();

			service.editAnyway();

			expect(service.passwordPromptOpen()).toBeTruthy();
			expect(service.lockActive()).toBeTruthy();
			expect(service.bannerActive()).toBeTruthy();
		});

		it('submitPassword with the correct password unlocks and closes the prompt', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			await service.submitPassword('right-password');

			expect(service.lockActive()).toBeFalsy();
			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.passwordError()).toBeNull();
			expect(service.bannerActive()).toBeFalsy();
		});

		it('submitPassword with a wrong password stays locked and reports wrong-password', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			await service.submitPassword('wrong-password');

			expect(service.lockActive()).toBeTruthy();
			expect(service.passwordPromptOpen()).toBeTruthy();
			expect(service.passwordError()).toBe('wrong-password');
		});

		it('cancelPasswordPrompt closes the prompt without unlocking', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			service.cancelPasswordPrompt();

			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.lockActive()).toBeTruthy();
			expect(service.bannerActive()).toBeTruthy();
		});

		it('ignores a password check that finishes after the next deck loads', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			const pending = service.submitPassword('right-password');
			service.resetForLoad();
			await pending;

			expect(service.lockActive()).toBeTruthy();
			expect(service.bannerActive()).toBeTruthy();
			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.checkingPassword()).toBeFalsy();
		});

		it('ignores a password check that finishes after Cancel', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			const pending = service.submitPassword('right-password');
			service.cancelPasswordPrompt();
			await pending;

			expect(service.lockActive()).toBeTruthy();
			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.passwordError()).toBeNull();
			expect(service.checkingPassword()).toBeFalsy();
		});

		it('resetForLoad re-arms the password prompt state for a newly loaded document', async () => {
			const verifier = await createModifyVerifier('right-password', { spinCount: 10 });
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();
			await service.submitPassword('wrong-password');
			expect(service.passwordError()).toBe('wrong-password');

			service.resetForLoad();

			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.passwordError()).toBeNull();
		});
	});

	describe('password-protected modifyVerifier with NO saltData (salt-less, still checkable)', () => {
		it('editAnyway opens the password prompt instead of unlocking', async () => {
			const verifier = await createSaltlessModifyVerifierForTesting('right-password', {
				spinCount: 10,
			});
			expect(verifier.saltData).toBeUndefined();
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			expect(service.recommendation().requiresPassword).toBeTruthy();

			service.editAnyway();

			expect(service.passwordPromptOpen()).toBeTruthy();
			expect(service.lockActive()).toBeTruthy();
			expect(service.bannerActive()).toBeTruthy();
		});

		it('submitPassword with the correct password unlocks and closes the prompt', async () => {
			const verifier = await createSaltlessModifyVerifierForTesting('right-password', {
				spinCount: 10,
			});
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			await service.submitPassword('right-password');

			expect(service.lockActive()).toBeFalsy();
			expect(service.passwordPromptOpen()).toBeFalsy();
			expect(service.bannerActive()).toBeFalsy();
		});

		it('submitPassword with a wrong password stays locked and reports wrong-password', async () => {
			const verifier = await createSaltlessModifyVerifierForTesting('right-password', {
				spinCount: 10,
			});
			const service = createService({ modifyVerifier: verifier } as unknown as PptxData);
			service.editAnyway();

			await service.submitPassword('wrong-password');

			expect(service.lockActive()).toBeTruthy();
			expect(service.passwordPromptOpen()).toBeTruthy();
			expect(service.passwordError()).toBe('wrong-password');
		});
	});
});

describe('loadNoticesService notify (failed Save)', () => {
	it('shows the save-failed notice first, again after a dismissal, until the next load', async () => {
		const { SAVE_FAILED_TOAST } = await import('ooxml-ui/pptx');
		const service = createService(undefined);
		service.notify(SAVE_FAILED_TOAST);
		expect(service.visibleToasts().map((toast) => toast.id)).toStrictEqual(['save-failed']);
		service.dismissToast('save-failed');
		expect(service.visibleToasts()).toStrictEqual([]);
		service.notify(SAVE_FAILED_TOAST);
		expect(service.visibleToasts().map((toast) => toast.id)).toStrictEqual(['save-failed']);
		service.resetForLoad();
		expect(service.visibleToasts()).toStrictEqual([]);
	});

	describe('signature load notice', () => {
		const invalid: SignatureSummary = {
			state: 'invalid',
			signatures: [{ path: '_xmlsignatures/sig1.xml', state: 'invalid' }],
		};
		function withSignatures(summary: SignatureSummary | undefined): LoadNoticesService {
			const injector = Injector.create({
				providers: [
					{
						provide: LoadContentService,
						useValue: {
							parsedData: signal(undefined),
							slides: signal([]),
							signatureSummary: signal(summary),
						},
					},
				],
			});
			return runInInjectionContext(injector, () => new LoadNoticesService());
		}

		it('tells a signed deck its state until dismissed, and again after the next load', () => {
			const service = withSignatures(invalid);
			const first = service.visibleToasts()[0];
			expect(first?.code).toBe('SIGNATURE_STATE');
			expect(first?.severity).toBe('warning');
			expect(first?.text).toContain('is not valid');
			expect(first?.text).toContain('Saving changes removes the signatures');
			service.dismissToast('SIGNATURE_STATE');
			expect(service.visibleToasts()).toHaveLength(0);
			service.resetForLoad();
			expect(service.visibleToasts()[0]?.code).toBe('SIGNATURE_STATE');
		});

		it('shows nothing for an unsigned deck', () => {
			expect(withSignatures(undefined).visibleToasts()).toHaveLength(0);
		});
	});
});
