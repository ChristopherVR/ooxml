// The Word source of the table cells fixture: a borderless table with wrapped and empty cells, a
// ruled table with bulleted and numbered lists in its cells, and a ruled table whose row breaks
// across the page. Word 16 exports it to PDF (export-word.ps1).
// Run from `src/core`: bun docx/pdf-import/__fixtures__/generate-tables.mjs
import { writeFile, mkdir } from 'node:fs/promises';
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = new Date('2026-11-24T00:00:00Z');
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const esc = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const run = (text, props = '') =>
	`<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const para = (runs, props = '') => `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ''}${runs}</w:p>`;
const plain = (text, props = '<w:spacing w:before="240" w:after="160"/>') => para(run(text), props);

const TEXT = [
	'Ce paragraphe sert de texte de remplissage générique, sans lien avec une organisation réelle.',
	'Il revient à la ligne plusieurs fois pour occuper la page avant le tableau suivant.',
];
const body = (count) => Array.from({ length: count }, (_, index) => TEXT[index % TEXT.length]).join(' ');

const STYLES = `${DECL}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="240"/><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>`;

const level = (ilvl, format, text, left, font = '') =>
	`<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>${font ? `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`;
const NUMBERING = `${DECL}<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${level(0, 'bullet', '\uf0b7', 360, 'Symbol')}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${level(0, 'decimal', '%1.', 360)}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

/** A cell: each `\n` starts a paragraph; `* ` and `# ` start bulleted and numbered items. */
const cell = (width, text, props = '') => {
	const paragraphs = text.split('\n').map((line) => {
		const list = /^([*#]) /.exec(line);
		if (!list) return para(line ? run(line) : '');
		const numId = list[1] === '*' ? 1 : 2;
		return para(run(line.slice(2)), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
	});
	return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${props}</w:tcPr>${paragraphs.join('')}</w:tc>`;
};
const table = (widths, rows, borders) =>
	`<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => (borders ? `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>` : `<w:${side} w:val="nil"/>`)).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${rows
		.map(
			(row, index) =>
				`<w:tr>${index === 0 && borders ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}${row.map((text, col) => cell(widths[col], text)).join('')}</w:tr>`,
		)
		.join('')}</w:tbl>`;

const plainWidths = [2000, 4000, 3000];
const plainRows = [
	['Élément', 'Description', 'Remarque'],
	['Papier', 'Feuilles blanches au format A4, achetées par ramettes de cinq cents feuilles.', 'Stock suffisant'],
	['Encre', '', 'À commander'],
	['Agrafes', 'Petite boîte', ''],
	['', 'Ligne sans élément', 'Note libre'],
	['Classeurs', 'Classeurs à levier, dos de huit centimètres, en carton recouvert.', 'Livrés en janvier, rangés dans l’armoire du couloir.'],
];

const listWidths = [2000, 3500, 3500];
const listRows = [
	['Étape', 'Actions', 'Contrôles'],
	['Préparer', '* Réunir les pièces.\n* Vérifier les dates.\n* Classer le dossier.', '# Relire la liste.\n# Signer le registre.'],
	['Envoyer', '* Imprimer le courrier.\n* Poster le pli.', '# Noter la date d’envoi.\n# Garder le reçu.\n# Archiver la copie.'],
];

const splitWidths = [1500, 5500, 2000];
const LONG = [
	'Première partie de la description, assez longue pour revenir à la ligne dans la cellule.',
	'Deuxième partie, qui continue la même ligne du tableau sur plusieurs lignes de texte.',
	'Troisième partie, encore une fois longue, pour que la ligne soit coupée par le saut de page.',
	'Quatrième partie de la description de cette ligne qui commence en bas de la première page.',
	'Cinquième partie, sur la page suivante, à la suite de la ligne coupée par le saut de page.',
	'Sixième et dernière partie de la description de la ligne coupée, avant la ligne suivante.',
];
const splitRows = [
	['Réf.', 'Description', 'Quantité'],
	...Array.from({ length: 4 }, (_, index) => [`R${index + 1}`, `Article ordinaire ${index + 1}`, String(10 + index)]),
	['R5', LONG.join('\n'), '50'],
	['R6', 'Article suivant', '60'],
	['R7', 'Dernier article', '70'],
];

const documentXml = [
	para(run('Tableaux et cellules', '<w:b/><w:sz w:val="32"/>'), '<w:pStyle w:val="Heading1"/>'),
	plain('Un tableau sans bordures, avec des cellules longues et des cellules vides :', '<w:spacing w:after="160"/>'),
	table(plainWidths, plainRows, false),
	plain('Un tableau avec des listes dans ses cellules :'),
	table(listWidths, listRows, true),
	plain(body(14), '<w:jc w:val="both"/><w:spacing w:before="240" w:after="160"/>'),
	plain('Un tableau dont une ligne est coupée par le saut de page :', '<w:spacing w:after="160"/>'),
	table(splitWidths, splitRows, true),
	plain('Fin des tableaux.'),
].join('');

const zip = new JSZip();
const add = (path, data) => zip.file(path, data, { date: DATE, createFolders: false });
add(
	'[Content_Types].xml',
	`${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${WML}.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="${WML}.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
);
add(
	'_rels/.rels',
	`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
);
add(
	'docProps/core.xml',
	`${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Tableaux et cellules</dc:title><dc:creator>Exemple</dc:creator></cp:coreProperties>`,
);
add(
	'word/_rels/document.xml.rels',
	`${DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
);
add('word/styles.xml', STYLES);
add('word/numbering.xml', NUMBERING);
add(
	'word/document.xml',
	`${DECL}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${documentXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
);
await mkdir(new URL('./source/', import.meta.url), { recursive: true });
await writeFile(
	new URL('./source/cellules.docx', import.meta.url),
	await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
);
