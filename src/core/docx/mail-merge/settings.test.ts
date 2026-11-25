import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRelationships } from '../../opc/index';
import { NS } from '../../xml/index';
import { createDocument } from '../model';
import { loadDocx } from '../parse';
import { saveDocx } from '../save';
import { paragraph, run, template } from './test-fixture';

const SETTINGS = 'word/settings.xml';
const RELS = 'word/_rels/settings.xml.rels';
const sourceType = `${NS.r}/mailMergeSource`;

describe('Word mail merge settings', () => {
	it('round-trips settings edits while preserving unmodeled children and unrelated relationships', async () => {
		const settings = `<w:settings xmlns:w="${NS.w}" xmlns:r="${NS.r}">
		<w:mailMerge><w:mainDocumentType w:val="formLetters"/><w:dataType w:val="spreadsheet"/>
		<w:query w:val="SELECT * FROM Sheet1"/><w:dataSource r:id="rId1"/>
		<w:viewMergedData/><w:odso><w:table w:val="Sheet1"/><w:src r:id="rId1"/>
		<w:fieldMapData><w:type w:val="dbColumn"/><w:name w:val="Given"/><w:mappedName w:val="FirstName"/><w:column w:val="0"/><w:lid w:val="en-US"/></w:fieldMapData>
		</w:odso></w:mailMerge><w:trackRevisions/></w:settings>`;
		const rels = `<Relationships xmlns="${NS.rels}">
		<Relationship Id="rId1" Type="${sourceType}" Target="file:///C:/old.xlsx" TargetMode="External"/>
		<Relationship Id="rId2" Type="${NS.r}/attachedTemplate" Target="letter.dotx" TargetMode="External"/>
		</Relationships>`;
		const loaded = await template(paragraph(run('Template')), {
			[SETTINGS]: settings,
			[RELS]: rels,
		});
		const base = loaded.model.mailMerge;
		expect(base).toMatchObject({
			dataSource: { relationshipId: 'rId1', target: 'file:///C:/old.xlsx', targetMode: 'External' },
			odsoSource: { relationshipId: 'rId1' },
			fieldMappings: [{ name: 'Given', mappedName: 'FirstName', column: 0, type: 'dbColumn' }],
		});
		expect(await loaded.save()).toEqual(loaded.bytes);
		const unchanged = await JSZip.loadAsync(
			await loaded.save({ ...loaded.model, page: { ...loaded.model.page, marginLeft: 100 } }),
		);
		expect(await unchanged.file(SETTINGS)?.async('string')).toBe(settings);
		expect(await unchanged.file(RELS)?.async('string')).toBe(rels);
		const bytes = await loaded.save({
			...loaded.model,
			mailMerge: {
				...base,
				dataSource: { target: 'file:///C:/new%20list.xlsx?x=1&y=2' },
				fieldMappings: [{ name: 'Forename', mappedName: 'FirstName', column: 2, type: 'dbColumn' }],
			},
		});
		const reopened = await loadDocx(bytes);
		expect(reopened.model.mailMerge).toMatchObject({
			dataSource: { target: 'file:///C:/new%20list.xlsx?x=1&y=2' },
			odsoSource: { target: 'file:///C:/old.xlsx' },
			fieldMappings: [{ name: 'Forename', column: 2 }],
		});
		const edited = await JSZip.loadAsync(bytes);
		expect(await edited.file(SETTINGS)?.async('string')).toContain('<w:lid w:val="en-US"/>');
		expect(await edited.file(SETTINGS)?.async('string')).toContain('<w:viewMergedData/>');
		expect(await edited.file(SETTINGS)?.async('string')).toContain('<w:table w:val="Sheet1"/>');
		const { mailMerge: _mailMerge, ...detached } = reopened.model;
		const finished = await JSZip.loadAsync(await reopened.save(detached));
		expect(await finished.file(SETTINGS)?.async('string')).toContain('<w:trackRevisions/>');
		expect(await finished.file(SETTINGS)?.async('string')).not.toContain('mailMerge');
		expect([...parseRelationships(await finished.file(RELS)?.async('string')).keys()]).toEqual([
			'rId2',
		]);
	});

	it('creates the settings part and source relationship in a new DOCX', async () => {
		const model = createDocument();
		model.mailMerge = { dataSource: { target: 'file:///C:/contacts.csv' } };
		const saved = await saveDocx(model);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('[Content_Types].xml')?.async('string')).toContain('/word/settings.xml');
		expect(await zip.file('word/_rels/document.xml.rels')?.async('string')).toContain(
			'relationships/settings',
		);
		expect((await loadDocx(saved)).model.mailMerge).toEqual({
			mainDocumentType: 'formLetters',
			dataType: 'textFile',
			dataSource: {
				relationshipId: 'rId1',
				target: 'file:///C:/contacts.csv',
				targetMode: 'External',
			},
		});
	});
});
