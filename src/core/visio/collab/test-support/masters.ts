// Helpers that add and edit masters in a package for the collaboration tests.
import type JSZip from 'jszip';

const NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MASTER_REL = 'http://schemas.microsoft.com/visio/2010/relationships/master';
const MASTERS_REL = 'http://schemas.microsoft.com/visio/2010/relationships/masters';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

export interface Added {
	id: string;
	unique: string;
	/** Id of a page shape that uses the master (none when omitted). */
	shapeId?: string;
	media?: { path: string; bytes: number[]; ext: string; type: string };
	/** The master part's shapes: the children of `Shapes` (default none). */
	shapes?: string;
}

/** Adds the first or a further master with a picture and a shape that uses it, as one local edit. */
export async function addMaster(zip: JSZip, added: Added): Promise<void> {
	const text = async (path: string) => (await zip.file(path)?.async('string')) ?? '';
	const dirRels = 'visio/masters/_rels/masters.xml.rels';
	const hasList = zip.file('visio/masters/masters.xml') !== null;
	const file = `visio/masters/master${added.id}.xml`;
	const entry = `<Master ID="${added.id}" NameU="M${added.id}" UniqueID="${added.unique}"><PageSheet/><Rel xmlns:r="${R}" r:id="rIdM${added.id}"/></Master>`;
	if (hasList) {
		const list = await text('visio/masters/masters.xml');
		zip.file('visio/masters/masters.xml', list.replace('</Masters>', `${entry}</Masters>`));
		zip.file(
			dirRels,
			(await text(dirRels)).replace(
				'</Relationships>',
				`<Relationship Id="rIdM${added.id}" Type="${MASTER_REL}" Target="master${added.id}.xml"/></Relationships>`,
			),
		);
	} else {
		zip.file('visio/masters/masters.xml', `<Masters xmlns="${NS}">${entry}</Masters>`);
		zip.file(
			dirRels,
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdM${added.id}" Type="${MASTER_REL}" Target="master${added.id}.xml"/></Relationships>`,
		);
		zip.file(
			'visio/_rels/document.xml.rels',
			(await text('visio/_rels/document.xml.rels')).replace(
				'</Relationships>',
				`<Relationship Id="rIdMasters" Type="${MASTERS_REL}" Target="masters/masters.xml"/></Relationships>`,
			),
		);
	}
	zip.file(
		file,
		`<MasterContents xmlns="${NS}">${added.shapes ? `<Shapes>${added.shapes}</Shapes>` : '<Shapes/>'}</MasterContents>`,
	);
	const media = added.media;
	if (media) {
		zip.file(
			`visio/masters/_rels/master${added.id}.xml.rels`,
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="${IMAGE_REL}" Target="../media/${media.path}"/></Relationships>`,
		);
		zip.file(`visio/media/${media.path}`, new Uint8Array(media.bytes));
	}
	let types = await text('[Content_Types].xml');
	if (!hasList)
		types = types.replace(
			'</Types>',
			'<Override PartName="/visio/masters/masters.xml" ContentType="application/vnd.ms-visio.masters+xml"/></Types>',
		);
	types = types.replace(
		'</Types>',
		`<Override PartName="/${file}" ContentType="application/vnd.ms-visio.master+xml"/>` +
			(!media || types.includes(`Extension="${media.ext}"`)
				? ''
				: `<Default Extension="${media.ext}" ContentType="${media.type}"/>`) +
			'</Types>',
	);
	zip.file('[Content_Types].xml', types);
	if (added.shapeId === undefined) return;
	const page = await text('visio/pages/page1.xml');
	zip.file(
		'visio/pages/page1.xml',
		page.replace(
			/<\/Shapes>/,
			`<Shape ID="${added.shapeId}" Type="Shape" Master="${added.id}"><Cell N="PinX" V="1"/></Shape></Shapes>`,
		),
	);
}

/** A master shape with a position, for `Added.shapes`. */
export const masterShape = (id: string, x: number, y = 1): string =>
	`<Shape ID="${id}" Type="Shape"><Cell N="PinX" V="${x}"/><Cell N="PinY" V="${y}"/></Shape>`;

/** Part paths of the master parts (not the list) of a package, sorted. */
export const masterPartPaths = (names: readonly string[]): string[] =>
	names.filter((n) => /^visio\/masters\/master\d+\.xml$/.test(n)).sort();

/** Sets the value of cell `name` of shape `shapeId` inside the package's only master part. */
export async function setMasterCell(
	zip: JSZip,
	shapeId: string,
	name: string,
	value: number,
): Promise<void> {
	const [path] = masterPartPaths(Object.keys(zip.files));
	const xml = (await zip.file(path!)!.async('string')) as string;
	const shape = new RegExp(`(<Shape ID="${shapeId}"[^>]*>[\\s\\S]*?<Cell N="${name}" V=")[^"]*(")`);
	if (!shape.test(xml)) throw new Error(`No cell ${name} on shape ${shapeId} of the master`);
	zip.file(path!, xml.replace(shape, `$1${value}$2`));
}

/** The value of cell `name` of shape `shapeId` inside master `masterId`'s part, from package text. */
export function masterCell(
	part: string | undefined,
	shapeId: string,
	name: string,
): string | undefined {
	return new RegExp(`<Shape ID="${shapeId}"[^>]*>[\\s\\S]*?<Cell N="${name}" V="([^"]*)"`).exec(
		part ?? '',
	)?.[1];
}

/** Removes master `id` (list entry, relationship, part and content type override). */
export async function removeMaster(zip: JSZip, id: string): Promise<void> {
	const list = 'visio/masters/masters.xml';
	const rels = 'visio/masters/_rels/masters.xml.rels';
	const text = async (path: string) => (await zip.file(path)?.async('string')) ?? '';
	zip.file(
		list,
		(await text(list)).replace(new RegExp(`<Master ID="${id}"[\\s\\S]*?</Master>`), ''),
	);
	zip.file(rels, (await text(rels)).replace(new RegExp(`<Relationship Id="rIdM${id}"[^>]*/>`), ''));
	zip.remove(`visio/masters/master${id}.xml`);
	zip.file(
		'[Content_Types].xml',
		(await text('[Content_Types].xml')).replace(
			new RegExp(`<Override PartName="/visio/masters/master${id}.xml"[^>]*/>`),
			'',
		),
	);
}
