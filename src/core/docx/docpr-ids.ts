// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Package-level allocator for `wp:docPr/@id`. Word expects drawing ids to be unique across the
// document, headers, footers and notes, so one allocator is shared by every part in a save.
import type JSZip from 'jszip';
import type { XmlDocument } from './xml';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DOC_PR_ID = /<(?:[\w.-]+:)?docPr\b[^>]*?\sid\s*=\s*["'](\d+)["']/g;
const OLE_SHAPE_ID = /<(?:[\w.-]+:)?shape\b[^>]*?\sid\s*=\s*["']_x0000_i(\d+)["']/g;
const VML_NS = 'urn:schemas-microsoft-com:vml';

export class DocPrIdAllocator {
	/**
	 * The ink that kept its own drawing (and so its part) in this save. Only their parts may be
	 * rewritten: a copy of ink still names the part it was copied from.
	 */
	readonly keptInk = new Set<object>();
	private max = 0;
	private used = new Set<number>();

	reserve(id: number): void {
		if (!Number.isSafeInteger(id) || id < 0) return;
		this.used.add(id);
		this.max = Math.max(this.max, id);
	}

	/** Reserves every `wp:docPr` id in an already parsed part. */
	reserveFromDocument(doc: XmlDocument): void {
		for (const node of Array.from(doc.getElementsByTagNameNS(WP_NS, 'docPr')))
			this.reserve(Number(node.getAttribute('id')));
		for (const node of Array.from(doc.getElementsByTagNameNS(VML_NS, 'shape'))) {
			const id = /^_x0000_i(\d+)$/.exec(node.getAttribute('id') ?? '');
			if (id) this.reserve(Number(id[1]) - 1024);
		}
	}

	/** Reserves every `docPr` id found in serialized XML. */
	reserveFromXml(xml: string): void {
		for (const match of xml.matchAll(DOC_PR_ID)) this.reserve(Number(match[1]));
		for (const match of xml.matchAll(OLE_SHAPE_ID)) this.reserve(Number(match[1]) - 1024);
	}

	/** Reserves the ids of every XML part in a package (document, headers, footers, notes...). */
	async reserveFromPackage(zip: JSZip): Promise<void> {
		for (const file of zip.file(/^word\/[^/]+\.xml$/))
			this.reserveFromXml(await file.async('string'));
	}

	next(preferred?: string): string {
		const id = preferred && /^\d+$/.test(preferred) ? Number(preferred) : NaN;
		if (Number.isInteger(id) && id <= 0xffffffff && !this.used.has(id)) {
			this.reserve(id);
			return String(id);
		}
		const next = this.max + 1;
		this.reserve(next);
		return String(next);
	}
}
