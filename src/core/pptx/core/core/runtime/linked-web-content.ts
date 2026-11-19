/**
 * Which relationships link a deck's pictures or media to a web address: the
 * content PowerPoint blocks by default and offers to download through its
 * security bar ("Enable Content"). A viewer needs to know a deck has some, so
 * it can offer the same choice after a load that left them out.
 *
 * @module runtime/linked-web-content
 */

/** Relationship types whose external target is downloaded to draw or play the deck. */
const LINKED_CONTENT_TYPES = [
	'/officeDocument/2006/relationships/image',
	'/officeDocument/2006/relationships/video',
	'/officeDocument/2006/relationships/audio',
	'/office/2007/relationships/media',
	'/office/2007/relationships/hdphoto',
];

/**
 * Whether a relationship links a picture or media file at an `http:` or
 * `https:` address. Hyperlinks, linked OLE objects and file paths are not
 * web content a load could fetch.
 */
export function isLinkedWebContent(type: string, target: string, targetMode: string): boolean {
	return (
		targetMode.toLowerCase() === 'external' &&
		/^https?:\/\//iu.test(target) &&
		LINKED_CONTENT_TYPES.some((suffix) => type.endsWith(suffix))
	);
}
