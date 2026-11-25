import { NS } from '../xml/index';
import { mapStoryBlocks } from './cell-blocks';
import type { Block, DocumentModel } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import { getW, importElement, WORD_NS, type XmlElement } from './xml';

export interface BlockCopy {
	sourceId: string;
	/** Shared by every block copied into the same recipient's document. */
	rekey(id: string): string;
}

export type BlockCopies = ReadonlyMap<string, BlockCopy>;

const copiesByDocument = new WeakMap<DocumentModel, BlockCopies>();

export function rememberBlockCopies(model: DocumentModel, copies: BlockCopies): void {
	if (copies.size) copiesByDocument.set(model, copies);
}

export function blockCopiesOf(model: DocumentModel): BlockCopies | undefined {
	return copiesByDocument.get(model);
}

const WORD_2010 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** Prepares a repeated body block for the normal paragraph/table writer. */
export function copySourceBlock(
	source: XmlElement,
	base: Block,
	copy: BlockCopy,
	allocator: RelationshipAllocator | undefined,
	bookmarkId: (id: string) => string,
): { node: XmlElement; base: Block } {
	const node = importElement(source.ownerDocument, source);
	for (const element of [node, ...Array.from(node.getElementsByTagName('*'))]) {
		if (element.namespaceURI === WORD_NS) {
			if (element.localName === 'p') {
				element.removeAttributeNS(WORD_2010, 'paraId');
				element.removeAttributeNS(WORD_2010, 'textId');
			}
			if (element.localName === 'bookmarkStart' || element.localName === 'bookmarkEnd') {
				const id = getW(element, 'id');
				if (id) element.setAttributeNS(WORD_NS, 'w:id', bookmarkId(copy.rekey(id)));
			}
		} else if (element.namespaceURI === NS.wp && element.localName === 'docPr' && allocator) {
			element.setAttribute('id', allocator.docPrIds.next());
		}
	}
	const rekey = (id: string) => copy.rekey(id);
	const [rekeyed] = mapStoryBlocks(
		[base],
		(paragraph) => ({ ...paragraph, id: rekey(paragraph.id) }),
		rekey,
	);
	return { node, base: rekeyed ?? base };
}
