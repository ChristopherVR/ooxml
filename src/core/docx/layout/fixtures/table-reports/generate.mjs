// Writes the table report fixtures (#320 to #323): generic content in the shape python-docx saves
// (Letter paper, 1" and 1.25" margins, "Table Grid", auto-fit tables). Run from `src/core`:
//   bun docx/layout/fixtures/table-reports/generate.mjs
import { writeFile } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const TEXT_WIDTH = 8640;

const escape = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const run = (text) => `<w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;
const paragraph = (text = '') => `<w:p>${text ? run(text) : ''}</w:p>`;
const line = (size) => `w:val="single" w:sz="${size}" w:space="0" w:color="000000"`;
const allBorders = (size) =>
	['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
		.map((side) => `<w:${side} ${line(size)}/>`)
		.join('');

/** A table as python-docx writes it: `tblW` auto, an even grid and a `tcW` per cell. */
function table(rows, options = {}) {
	const columns = rows[0].length;
	const widths = options.widths ?? Array.from({ length: columns }, () => TEXT_WIDTH / columns);
	const tblW = options.width
		? `<w:tblW w:w="${options.width}" w:type="dxa"/>`
		: '<w:tblW w:w="0" w:type="auto"/>';
	const properties = [
		'<w:tblStyle w:val="TableGrid"/>',
		tblW,
		options.center ? '<w:jc w:val="center"/>' : '',
		options.borders ? `<w:tblBorders>${allBorders(options.borders)}</w:tblBorders>` : '',
		options.cellMargin
			? `<w:tblCellMar><w:top w:w="${options.cellMargin}" w:type="dxa"/><w:left w:w="${options.cellMargin}" w:type="dxa"/><w:bottom w:w="${options.cellMargin}" w:type="dxa"/><w:right w:w="${options.cellMargin}" w:type="dxa"/></w:tblCellMar>`
			: '',
		'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
	].join('');
	const body = rows
		.map((cells) => {
			const trPr = options.cantSplit ? '<w:trPr><w:cantSplit/></w:trPr>' : '';
			const tcs = cells
				.map((content, index) => {
					const paragraphs = Array.isArray(content) ? content : [content];
					return `<w:tc><w:tcPr><w:tcW w:w="${widths[index]}" w:type="dxa"/></w:tcPr>${paragraphs.map(paragraph).join('')}</w:tc>`;
				})
				.join('');
			return `<w:tr>${trPr}${tcs}</w:tr>`;
		})
		.join('');
	return `<w:tbl><w:tblPr>${properties}</w:tblPr><w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${body}</w:tbl>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>${allBorders(4)}</w:tblBorders></w:tblPr></w:style></w:styles>`;

async function save(name, body, stories = {}) {
	const zip = new JSZip();
	// A fixed date keeps the archive bytes stable between runs.
	const add = (part, text) => zip.file(part, text, { date: new Date('2026-11-24T00:00:00Z') });
	const overrides = [
		['/word/document.xml', `${MAIN}.document.main+xml`],
		['/word/styles.xml', `${MAIN}.styles+xml`],
		...(stories.header ? [['/word/header1.xml', `${MAIN}.header+xml`]] : []),
		...(stories.footer ? [['/word/footer1.xml', `${MAIN}.footer+xml`]] : []),
	];
	add(
		'[Content_Types].xml',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')}</Types>`,
	);
	add(
		'_rels/.rels',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	const rels = [
		`<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/>`,
		stories.header ? `<Relationship Id="rId2" Type="${R}/header" Target="header1.xml"/>` : '',
		stories.footer ? `<Relationship Id="rId3" Type="${R}/footer" Target="footer1.xml"/>` : '',
	].join('');
	add(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG}">${rels}</Relationships>`,
	);
	add('word/styles.xml', STYLES);
	if (stories.header)
		add(
			'word/header1.xml',
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}" xmlns:r="${R}">${stories.header}</w:hdr>`,
		);
	if (stories.footer)
		add(
			'word/footer1.xml',
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="${W}" xmlns:r="${R}">${stories.footer}</w:ftr>`,
		);
	const references = [
		stories.header ? '<w:headerReference w:type="default" r:id="rId2"/>' : '',
		stories.footer ? '<w:footerReference w:type="default" r:id="rId3"/>' : '',
	].join('');
	const sectPr = `<w:sectPr>${references}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
	add(
		'word/document.xml',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${sectPr}</w:body></w:document>`,
	);
	const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
	await writeFile(new URL(`./${name}.docx`, import.meta.url), bytes);
}

const columns12 = Array.from({ length: 12 }, (_, index) => `Colonne${index + 1}`);
await save(
	't02-autofit-12-colonnes',
	paragraph('Tableau auto-ajusté de douze colonnes.') +
		table([
			columns12,
			...[1, 2, 3].map((row) => columns12.map((_, column) => `Val${row}.${column + 1}`)),
		]) +
		paragraph('Fin du document.'),
);

await save(
	't03-mots-insecables',
	paragraph('Coordonnées bancaires et références sans espaces.') +
		table([
			['Libellé', 'Valeur'],
			[
				'Adresse',
				'https://www.example.com/documents/2026/rapport-annuel/annexe-technique/version-finale-definitive.pdf',
			],
			['IBAN', 'FR7630006000011234567890189'.repeat(3)],
			['Référence', 'REF-2026-000123456789-ABCDEFGHIJKLMNOPQRSTUVWXYZ-0123456789'],
		]) +
		paragraph('Fin du document.'),
);

await save(
	't10-tableau-entete-pied',
	paragraph('Corps de la lettre.') +
		table([
			['Objet', 'Dossier'],
			['Contrat', 'Signé'],
		]) +
		paragraph('Fin du document.'),
	{
		header:
			table([
				['Notre référence : DOS-2026-001', 'Votre référence : CLI-42'],
				['Affaire : Exemple contre Exemple', 'Date : 24 novembre 2026'],
			]) + paragraph(),
		footer:
			table([
				['Cabinet Exemple', 'Société d’avocats', 'Toque P 000'],
				['1 rue de l’Exemple', '75000 Paris', 'contact@example.com'],
			]) + paragraph(),
	},
);

await save(
	't14-ligne-plus-haute-que-page',
	paragraph('Paragraphe avant le tableau.') +
		table([[Array.from({ length: 80 }, (_, index) => `Paragraphe ${index + 1}`)]], {
			cantSplit: true,
		}) +
		paragraph('Paragraphe après le tableau.'),
);

await save(
	't15-centre-bordures-epaisses',
	paragraph('Tableau centré, bordures épaisses et grandes marges de cellule.') +
		table(
			[
				['Désignation', 'Quantité', 'Montant'],
				['Prestation de conseil', '12', '3 600,00 €'],
				['Frais de dossier', '1', '150,00 €'],
			],
			{ center: true, borders: 24, cellMargin: 340, width: TEXT_WIDTH },
		) +
		paragraph('Fin du document.'),
);
