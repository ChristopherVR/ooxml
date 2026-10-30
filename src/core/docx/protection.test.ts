import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from './index';
import { expectParagraph } from './test-support/access';
import { enforceProtection, verifyProtectionPassword } from './protection';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;

async function load(body: string, settings = '') {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	zip.file('word/settings.xml', `<w:settings xmlns:w="${W}">${settings}</w:settings>`);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const part = async (bytes: Uint8Array, name: string) =>
	(await (await JSZip.loadAsync(bytes)).file(name)?.async('string')) ?? '';

describe('protection passwords', () => {
	it('checks a password against the verifier it created, and refuses a wrong one', () => {
		const protection = enforceProtection('readOnly', { password: 'Secret1', salt: 'AAECAwQFBgc=' });
		expect(protection.password).toMatchObject({
			algorithmName: 'SHA-512',
			spinCount: 100000,
			word: { algorithmSid: 14 },
		});
		expect(verifyProtectionPassword(protection, 'Secret1')).toBe(true);
		expect(verifyProtectionPassword(protection, 'secret1')).toBe(false);
		expect(verifyProtectionPassword(protection, '')).toBe(false);
	});

	it('lets protection without a password stop with any entry', () => {
		expect(verifyProtectionPassword(enforceProtection('comments'), '')).toBe(true);
		expect(verifyProtectionPassword(undefined, 'x')).toBe(true);
	});

	it('reads both attribute spellings and verifies them', async () => {
		const made = enforceProtection('readOnly', { password: 'pw', salt: 'AAECAwQFBgc=' }).password!;
		const word = await load(
			text('A'),
			`<w:documentProtection w:edit="readOnly" w:enforcement="1" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="${made.hashValue}" w:salt="${made.saltValue}"/>`,
		);
		expect(verifyProtectionPassword(word.model.protection, 'pw')).toBe(true);
		expect(verifyProtectionPassword(word.model.protection, 'no')).toBe(false);
		const iso = await load(
			text('A'),
			`<w:documentProtection w:edit="comments" w:enforcement="1" w:algorithmName="SHA-512" w:hashValue="${made.hashValue}" w:saltValue="${made.saltValue}" w:spinCount="100000"/>`,
		);
		expect(iso.model.protection).toMatchObject({ edit: 'comments', enforced: true });
		expect(verifyProtectionPassword(iso.model.protection, 'pw')).toBe(true);
	});

	it('never unlocks a hash it cannot check', async () => {
		const loaded = await load(
			text('A'),
			'<w:documentProtection w:edit="readOnly" w:enforcement="1" w:hash="ABCD"/>',
		);
		expect(loaded.model.protection?.unverifiablePassword).toBe(true);
		expect(verifyProtectionPassword(loaded.model.protection, 'anything')).toBeUndefined();
	});
});

describe('protection round trip', () => {
	it('writes protection started in the editor and reads it back', async () => {
		const loaded = await load(`<w:p>${text('A')}</w:p>`, '<w:trackRevisions/>');
		loaded.model.protection = enforceProtection('trackedChanges', {
			password: 'pw',
			formatting: true,
			salt: 'AAECAwQFBgc=',
		});
		const saved = await loaded.save();
		const settings = await part(saved, 'word/settings.xml');
		expect(settings).toMatch(/<w:documentProtection w:edit="trackedChanges" w:formatting="1"/);
		expect(settings.indexOf('trackRevisions')).toBeLessThan(settings.indexOf('documentProtection'));
		const again = await loadDocx(saved);
		expect(again.model.protection).toEqual(loaded.model.protection);
		expect(verifyProtectionPassword(again.model.protection, 'pw')).toBe(true);
	});

	it('removes the element when the protection is stopped, and leaves an untouched one alone', async () => {
		const source = '<w:documentProtection w:edit="readOnly" w:enforcement="1" w:hash="ABCD"/>';
		const loaded = await load(`<w:p>${text('A')}</w:p>`, source);
		expect(await part(await loaded.save(), 'word/settings.xml')).toContain('w:hash="ABCD"');
		delete loaded.model.protection;
		expect(await part(await loaded.save(), 'word/settings.xml')).not.toContain(
			'documentProtection',
		);
	});
});

describe('editable ranges', () => {
	const start = (id: number, group = 'everyone') =>
		`<w:permStart w:id="${id}" w:edGrp="${group}"/>`;
	const end = (id: number) => `<w:permEnd w:id="${id}"/>`;

	it('reads a range inside a paragraph, and one that spans paragraphs and table cells', async () => {
		const loaded = await load(
			`<w:p>${text('ab')}${start(1)}${text('cd')}${end(1)}${text('ef')}</w:p>` +
				`<w:p>${text('xx')}${start(2)}${text('yy')}</w:p><w:p>${text('zz')}</w:p>` +
				`<w:p>${text('qq')}${end(2)}${text('rr')}</w:p>`,
		);
		const permissions = (index: number) => expectParagraph(loaded.model.blocks[index]).permissions;
		expect(permissions(0)).toEqual([{ id: '1', group: 'everyone', from: 2, to: 4 }]);
		expect(permissions(1)).toEqual([{ id: '2', group: 'everyone', from: 2, to: 4, toNext: true }]);
		expect(permissions(2)).toEqual([
			{ id: '2', group: 'everyone', from: 0, to: 2, fromPrevious: true, toNext: true },
		]);
		expect(permissions(3)).toEqual([
			{ id: '2', group: 'everyone', from: 0, to: 2, fromPrevious: true },
		]);
	});

	it('keeps the markers through an edit and writes ranges the pane changed', async () => {
		const loaded = await load(`<w:p>${text('ab')}${start(1)}${text('cd')}${end(1)}</w:p>`);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs[0]!.text = 'abc';
		const kept = await part(await loaded.save(), 'word/document.xml');
		expect(kept).toMatch(/<w:permStart[^>]*w:id="1"[^>]*\/>.*<w:permEnd[^>]*w:id="1"/s);
		paragraph.permissions = [{ id: '7', group: 'everyone', from: 0, to: 2 }];
		const written = await part(await loaded.save(), 'word/document.xml');
		expect(written).not.toContain('w:id="1"');
		const again = expectParagraph((await loadDocx(await loaded.save())).model.blocks[0]);
		expect(again.permissions).toEqual([{ id: '7', group: 'everyone', from: 0, to: 2 }]);
		delete paragraph.permissions;
		expect(await part(await loaded.save(), 'word/document.xml')).not.toContain('permStart');
	});
});
