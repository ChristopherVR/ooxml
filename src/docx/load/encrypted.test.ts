import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { decryptOoxmlPackage } from '../../crypto/index.js';
import { createDocument, saveDocx } from '../index.js';
import { detectDocumentFormat, loadDocument } from './detect.js';
import { IncorrectPasswordError, PasswordRequiredError, encryptOoxmlPackage } from './index.js';
import { LegacyDocError, loadLegacyDoc } from './legacy-doc.js';

// word-encrypted.docx: saved by Word 16 with the password to open 'open sesame'
// (generate-word-encrypted.ps1).
const fixture = async (): Promise<Uint8Array> =>
	new Uint8Array(await readFile(new URL('./__fixtures__/word-encrypted.docx', import.meta.url)));
const PASSWORD = 'open sesame';

const text = (blocks: unknown[]): string => JSON.stringify(blocks);

describe('encrypted documents', () => {
	it('detects a Word-encrypted .docx as encrypted, not as a legacy .doc', async () => {
		const bytes = await fixture();
		expect(detectDocumentFormat(bytes)).toBe('encrypted');
		await expect(loadLegacyDoc(bytes)).rejects.toBeInstanceOf(LegacyDocError);
	});

	it('requires the password and rejects a wrong one with typed errors', async () => {
		const missing = await loadDocument(await fixture()).catch((error: unknown) => error);
		expect(missing).toBeInstanceOf(PasswordRequiredError);
		expect((missing as PasswordRequiredError).code).toBe('password-required');
		const wrong = await loadDocument(await fixture(), { password: 'nope' }).catch(
			(error: unknown) => error,
		);
		expect(wrong).toBeInstanceOf(IncorrectPasswordError);
		expect((wrong as IncorrectPasswordError).code).toBe('incorrect-password');
	}, 60_000);

	it('opens the Word-encrypted document with its password', async () => {
		const loaded = await loadDocument(await fixture(), { password: PASSWORD });
		expect(text(loaded.model.blocks)).toContain('Top secret paragraph.');
		expect(loaded.model.warnings.some((warning) => /password/.test(warning))).toBe(true);
		const zip = await JSZip.loadAsync(await decryptOoxmlPackage(await fixture(), PASSWORD));
		const core = (await zip.file('docProps/core.xml')?.async('string')) ?? '';
		expect(core).not.toMatch(/<dc:creator>[^<]+<\/dc:creator>/);
		expect(core).not.toMatch(/<cp:lastModifiedBy>[^<]+<\/cp:lastModifiedBy>/);
	}, 60_000);

	it('round-trips a document encrypted with our own writer', async () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Ours, encrypted.' }] }];
		const encrypted = await encryptOoxmlPackage(await saveDocx(model), PASSWORD, {
			spinCount: 1000,
		});
		expect(detectDocumentFormat(encrypted)).toBe('encrypted');
		const loaded = await loadDocument(encrypted, { password: PASSWORD });
		expect(text(loaded.model.blocks)).toContain('Ours, encrypted.');
	});
});
