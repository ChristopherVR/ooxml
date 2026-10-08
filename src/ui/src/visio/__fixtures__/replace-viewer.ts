import JSZip from 'jszip';
import { createVsdxFixture } from './fixture.mjs';

/** Independent synthetic pages; repeated local IDs exercise page-qualified occurrence identities. */
export async function replaceFixture(protectedSecond = false, richSecond = false) {
	const zip = await JSZip.loadAsync(await createVsdxFixture('cat cat CAT'));
	const first = await zip.file('visio/pages/page1.xml')!.async('string');
	const original = first.match(/<Shape\b.*?<\/Shape>/su)![0];
	const second = original
		.replace('ID="1"', 'ID="2"')
		.replace('cat cat CAT', richSecond ? '<cp IX="0"/>cat <cp IX="1"/>cat' : 'cat Ω cat')
		.replace(
			'<Text>',
			`${protectedSecond ? '<Cell N="LockTextEdit" V="1"/>' : ''}${richSecond ? '<Section N="Character"><Row IX="0"><Cell N="Style" V="0"/></Row><Row IX="1"><Cell N="Style" V="1"/></Row></Section>' : ''}<Text>`,
		);
	zip.file('visio/pages/page1.xml', first.replace('</Shapes>', `${second}</Shapes>`));
	zip.file('visio/pages/page2.xml', first.replace('cat cat CAT', 'cat tail'));
	const pages = await zip.file('visio/pages/pages.xml')!.async('string');
	const page2 = pages
		.match(/<Page ID="1".*?<\/Page>/su)![0]
		.replace('ID="1"', 'ID="2"')
		.replace('Imported page', 'Second page')
		.replace('rId1', 'rId2');
	zip.file('visio/pages/pages.xml', pages.replace('</Pages>', `${page2}</Pages>`));
	const rels = await zip.file('visio/pages/_rels/pages.xml.rels')!.async('string');
	zip.file(
		'visio/pages/_rels/pages.xml.rels',
		rels.replace(
			'</Relationships>',
			'<Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page2.xml"/></Relationships>',
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}
