// Paragraph content: formatted text, breaks, pictures, fields, bookmarks, note and comment
// references and anchored shapes, as WordprocessingML runs.
import type { DocInline } from '@christophervr/ole2/doc-model-types';
import { RELATIONSHIP_TYPES } from '../../../opc/index';
import type { DocCharProps } from '@christophervr/ole2/doc-char-props';
import { floatingPictureXml, inlinePictureXml, textBoxXml, type StoryWriter } from './drawing';
import { runPropertiesXml, wrapRunProperties } from './run-props';
import { attrs, el, esc } from './xml';

const SPECIAL_CHARACTERS: Record<string, string> = {
	'\t': '<w:tab/>',
	'‑': '<w:noBreakHyphen/>',
	'­': '<w:softHyphen/>',
};

/** `<w:t>`, `<w:tab/>` and the hyphen elements for a stretch of text. */
export function textContentXml(text: string): string {
	let out = '';
	let plain = '';
	const flush = () => {
		if (plain) out += `<w:t xml:space="preserve">${esc(plain)}</w:t>`;
		plain = '';
	};
	for (const character of text) {
		const special = SPECIAL_CHARACTERS[character];
		if (special) {
			flush();
			out += special;
		} else plain += character;
	}
	flush();
	return out;
}

class InlineWriter {
	constructor(
		private readonly writer: StoryWriter,
		private readonly paragraphStyle: number | undefined,
	) {}

	private get conversion() {
		return this.writer.conversion;
	}

	private run(props: DocCharProps | undefined, content: string): string {
		if (!content) return '';
		const base = this.conversion.charBase(this.paragraphStyle, props?.styleId);
		const rPr = props ? runPropertiesXml(props, base, this.conversion) : '';
		return `<w:r>${wrapRunProperties(rPr)}${content}</w:r>`;
	}

	items(items: readonly DocInline[]): string {
		let out = '';
		for (let i = 0; i < items.length; i++) {
			const item = items[i]!;
			if (item.type === 'fieldBegin') {
				const link = this.hyperlink(items, i);
				if (link) {
					out += link.xml;
					i = link.end;
					continue;
				}
			}
			if (item.type === 'fieldInstruction') {
				// Word splits an instruction around nested fields; one run keeps the code readable.
				let text = item.text;
				for (
					let next = items[i + 1];
					next?.type === 'fieldInstruction' && next.id === item.id;
					next = items[i + 1]
				)
					text += (items[++i] as typeof item).text;
				out += this.item({ ...item, text });
				continue;
			}
			out += this.item(item);
		}
		return out;
	}

	private hyperlink(
		items: readonly DocInline[],
		begin: number,
	): { xml: string; end: number } | undefined {
		const start = items[begin];
		if (start?.type !== 'fieldBegin' || start.field?.kind !== 'hyperlink') return undefined;
		const { url, anchor, tooltip } = start.field;
		if (!url && !anchor) return undefined;
		let separate = -1;
		let end = -1;
		for (let j = begin + 1; j < items.length; j++) {
			const item = items[j]!;
			if (item.type === 'fieldSeparate' && item.id === start.id && separate < 0) separate = j;
			if (item.type === 'fieldEnd' && item.id === start.id) {
				end = j;
				break;
			}
		}
		if (separate < 0 || end < 0) return undefined;
		const relationship = url
			? this.conversion.rels.add(RELATIONSHIP_TYPES.hyperlink, url, true)
			: undefined;
		const open = `<w:hyperlink${attrs({
			'r:id': relationship,
			'w:anchor': anchor,
			'w:tooltip': tooltip,
			'w:history': '1',
		})}>`;
		return { xml: `${open}${this.items(items.slice(separate + 1, end))}</w:hyperlink>`, end };
	}

	private item(item: DocInline): string {
		const { conversion } = this;
		switch (item.type) {
			case 'text':
				return this.run(item.props, textContentXml(item.text));
			case 'break':
				return this.run(
					item.props,
					item.kind === 'line' ? '<w:br/>' : `<w:br w:type="${item.kind}"/>`,
				);
			case 'picture':
				return this.run(item.props, inlinePictureXml(conversion, item.picture));
			case 'ole': {
				const progId = item.object.progId ?? 'an embedded object';
				conversion.warn(
					`Embedded OLE objects are shown as their preview picture; the object data (${progId}) is not kept.`,
				);
				return item.object.preview
					? this.run(item.props, inlinePictureXml(conversion, item.object.preview, 'Object'))
					: '';
			}
			case 'fieldBegin':
				return this.run(undefined, '<w:fldChar w:fldCharType="begin"/>');
			case 'fieldInstruction':
				return this.run(
					item.props,
					`<w:instrText xml:space="preserve">${esc(item.text)}</w:instrText>`,
				);
			case 'fieldSeparate':
				return this.run(undefined, '<w:fldChar w:fldCharType="separate"/>');
			case 'fieldEnd':
				return this.run(undefined, '<w:fldChar w:fldCharType="end"/>');
			case 'bookmarkStart':
				return item.bookmark.name
					? el('w:bookmarkStart', { 'w:id': item.id, 'w:name': item.bookmark.name })
					: '';
			case 'bookmarkEnd':
				return item.bookmark.name ? el('w:bookmarkEnd', { 'w:id': item.id }) : '';
			case 'noteReference':
				return this.run(
					item.props,
					el(item.kind === 'footnote' ? 'w:footnoteReference' : 'w:endnoteReference', {
						'w:id': item.id + 1,
					}),
				);
			case 'noteMark':
				return this.run(
					item.props,
					item.kind === 'comment'
						? '<w:annotationRef/>'
						: item.kind === 'footnote'
							? '<w:footnoteRef/>'
							: '<w:endnoteRef/>',
				);
			case 'commentReference':
				// The .doc comment range is not decoded, so the comment is anchored at its mark.
				conversion.warn(
					'Comments are anchored at their reference mark; the commented text range is not recovered from the .doc file.',
				);
				return (
					el('w:commentRangeStart', { 'w:id': item.id }) +
					el('w:commentRangeEnd', { 'w:id': item.id }) +
					this.run(item.props, el('w:commentReference', { 'w:id': item.id }))
				);
			case 'separator':
				return '';
			case 'shapeAnchor':
				if (item.shape.isPicture)
					return this.run(item.props, floatingPictureXml(conversion, item.shape));
				conversion.warn(
					'Drawing shapes other than pictures and text boxes (lines, rectangles, groups) are not converted from .doc files.',
				);
				return '';
			case 'textBoxAnchor':
				return this.run(
					item.props,
					textBoxXml(this.writer, item.shape, conversion.model.textBoxes[item.textBox]),
				);
		}
	}
}

/** The runs of a paragraph's inline items. */
export function inlinesXml(
	writer: StoryWriter,
	items: readonly DocInline[],
	paragraphStyle: number | undefined,
): string {
	return new InlineWriter(writer, paragraphStyle).items(items);
}
