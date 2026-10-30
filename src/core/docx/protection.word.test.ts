// Documents protected by desktop Word itself (generate-word-protection.ps1).
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { verifyProtectionPassword } from './protection';

const WORD_PASSWORD = 'Word-Pass-1';
const ISO_PASSWORD = 'Iso-Pass-1';
const fixture = (name: string) =>
	new Uint8Array(readFileSync(new URL(`./__fixtures__/protection/${name}.docx`, import.meta.url)));

describe('documents protected by Word', () => {
	for (const [name, edit] of [
		['readOnly', 'readOnly'],
		['comments', 'comments'],
		['trackedChanges', 'trackedChanges'],
		['forms', 'forms'],
		['readOnly-compat12', 'readOnly'],
	] as const) {
		it(`${name}: is read as enforced Word protection that is kept but not checked`, async () => {
			const { model } = await loadDocx(fixture(name));
			expect(model.protection).toMatchObject({ edit, enforced: true });
			expect(model.protection?.password?.word?.algorithmSid).toBe(14);
			expect(model.protection?.password?.algorithmName).toBe('SHA-512');
			// Word's own hash includes a legacy key step this code does not implement: Word accepts
			// "Word-Pass-1", and the plain spin of it does not reproduce the stored hash.
			expect(verifyProtectionPassword(model.protection, WORD_PASSWORD)).toBeUndefined();
		});
	}

	it('writes Word protection back as Word wrote it', async () => {
		const loaded = await loadDocx(fixture('readOnly'));
		const again = await loadDocx(await loaded.save());
		expect(again.model.protection).toEqual(loaded.model.protection);
	});

	it('verifies the ISO 29500 form, which Word itself unprotects with the right password', async () => {
		// readOnly-iso.docx: generate-word-protection.ps1 rewrites Word's element to
		// algorithmName/hashValue/saltValue/spinCount and checks Document.Unprotect accepts it.
		const { model } = await loadDocx(fixture('readOnly-iso'));
		expect(model.protection?.password?.word).toBeUndefined();
		expect(verifyProtectionPassword(model.protection, ISO_PASSWORD)).toBe(true);
		expect(verifyProtectionPassword(model.protection, 'iso-pass-1')).toBe(false);
		expect(verifyProtectionPassword(model.protection, WORD_PASSWORD)).toBe(false);
	});
});

describe('ISO 29500 digests', () => {
	/** The password spin written out with node:crypto, independent of the digest area. */
	const spin = (algorithm: string, password: string, salt: Buffer, count: number) => {
		let hash = createHash(algorithm)
			.update(Buffer.concat([salt, Buffer.from(password, 'utf16le')]))
			.digest();
		for (let i = 0; i < count; i++) {
			const iterator = Buffer.alloc(4);
			iterator.writeUInt32LE(i);
			hash = createHash(algorithm)
				.update(Buffer.concat([hash, iterator]))
				.digest();
		}
		return hash.toString('base64');
	};
	const salt = Buffer.from('0123456789abcdef');

	for (const [name, algorithm] of [
		['SHA-1', 'sha1'],
		['SHA-256', 'sha256'],
		['SHA-384', 'sha384'],
		['SHA-512', 'sha512'],
	] as const) {
		it(`checks ${name}`, async () => {
			const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
			const zip = new JSZip();
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${W}"><w:body><w:p/><w:sectPr/></w:body></w:document>`,
			);
			zip.file(
				'word/settings.xml',
				`<w:settings xmlns:w="${W}"><w:documentProtection w:edit="readOnly" w:enforcement="1" w:algorithmName="${name}" w:hashValue="${spin(algorithm, 'pässword', salt, 5000)}" w:saltValue="${salt.toString('base64')}" w:spinCount="5000"/></w:settings>`,
			);
			const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			expect(verifyProtectionPassword(model.protection, 'pässword')).toBe(true);
			expect(verifyProtectionPassword(model.protection, 'password')).toBe(false);
		});
	}
});
