// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { PendingMediaPart } from './model';
import type { NumberingCatalog, NumberingPictureBullet } from './numbering-model';
import { NUMBERING_PART } from './numbering-picture';
import { parseRelationships } from './package-parts';
import { relationshipsPartFor, writeNewRelationships } from './part-relationships';
import { makeNS, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml';

const VML_NS = 'urn:schemas-microsoft-com:vml';
const OFFICE_NS = 'urn:schemas-microsoft-com:office:office';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_REL = `${R_NS}/image`;

/** The picture bullets `next` adds to `prior`; throws when an existing one was changed or removed. */
export function addedPictureBullets(
	prior: NumberingCatalog | undefined,
	next: NumberingCatalog,
): NumberingPictureBullet[] {
	for (const [id, bullet] of Object.entries(prior?.pictureBullets ?? {}))
		if (JSON.stringify(next.pictureBullets?.[id]) !== JSON.stringify(bullet))
			throw new Error(
				`Cannot edit picture bullet "${id}"; the original numbering.xml and its images are preserved unchanged.`,
			);
	return Object.values(next.pictureBullets ?? {}).filter(
		(bullet) => !prior?.pictureBullets?.[bullet.id],
	);
}

function buildPictureBullet(
	doc: XmlDocument,
	bullet: NumberingPictureBullet,
	relId: string,
	serial: number,
): XmlElement {
	const element = makeW(doc, 'numPicBullet');
	element.setAttributeNS(WORD_NS, 'w:numPicBulletId', bullet.id);
	const pict = makeW(doc, 'pict');
	const shape = makeNS(doc, VML_NS, 'v:shape');
	shape.setAttribute('id', `_x0000_i${1025 + serial}`);
	shape.setAttribute('type', '#_x0000_t75');
	const width = bullet.widthPt ?? bullet.heightPt ?? 9;
	const height = bullet.heightPt ?? width;
	shape.setAttribute('style', `width:${+width.toFixed(2)}pt;height:${+height.toFixed(2)}pt`);
	shape.setAttributeNS(OFFICE_NS, 'o:bullet', 't');
	const imagedata = makeNS(doc, VML_NS, 'v:imagedata');
	imagedata.setAttributeNS(R_NS, 'r:id', relId);
	imagedata.setAttributeNS(OFFICE_NS, 'o:title', '');
	shape.appendChild(imagedata);
	pict.appendChild(shape);
	element.appendChild(pict);
	return element;
}

/**
 * Writes new `numPicBullet` entries before `before` (the first definition), with their image
 * parts and the relationships in `word/_rels/numbering.xml.rels`. Existing entries and
 * relationships are kept. An image already in the package is only referenced.
 */
export async function writePictureBullets(
	zip: JSZip,
	doc: XmlDocument,
	added: readonly NumberingPictureBullet[],
	before: XmlElement | null,
	pendingMedia: ReadonlyMap<string, PendingMediaPart> | undefined,
): Promise<void> {
	if (!added.length) return;
	const used = new Set(
		parseRelationships(
			await zip.file(relationshipsPartFor(NUMBERING_PART))?.async('string'),
		).keys(),
	);
	let serial = 1;
	const relationships = added.map((bullet, index) => {
		while (used.has(`rId${serial}`)) serial++;
		const id = `rId${serial}`;
		used.add(id);
		doc.documentElement.insertBefore(buildPictureBullet(doc, bullet, id, index), before);
		const stored = zip.file(bullet.partName) !== null;
		if (!stored && !pendingMedia?.get(bullet.partName))
			throw new Error(`Missing bytes for picture bullet image: ${bullet.partName}`);
		return {
			id,
			type: IMAGE_REL,
			target: bullet.partName.replace(/^word\//, ''),
			mode: 'Internal' as const,
			...(stored ? {} : { partName: bullet.partName }),
		};
	});
	await writeNewRelationships(zip, NUMBERING_PART, relationships, pendingMedia);
}
