// Block content: paragraphs (style, properties, runs, section break) and tables, for every story
// (body, header, footer, note, comment, text box).
import type {
	DocBlock,
	DocModelStory,
	DocParagraphBlock,
	DocTableBlock,
} from '@christophervr/ole2/doc-model-types';
import type { Conversion } from './context';
import type { StoryWriter } from './drawing';
import { inlinesXml } from './inline';
import { paragraphPropertiesXml } from './para-props';
import { runPropertiesXml, wrapRunProperties } from './run-props';
import { tableXml } from './table';
import { el } from './xml';

const MAX_NESTED_TEXT_BOXES = 2;

export class BlockWriter implements StoryWriter {
	private textBoxDepth = 0;
	constructor(
		readonly conversion: Conversion,
		/** The `w:sectPr` of the section a paragraph mark ends. */
		private readonly sectionProperties: (section: number) => string,
	) {}

	paragraph(block: DocParagraphBlock): string {
		const { conversion } = this;
		const styleId = this.paragraphStyleId(block);
		let pPr = styleId ? el('w:pStyle', { 'w:val': styleId }) : '';
		pPr += paragraphPropertiesXml(block.props, conversion.styles.paragraph(block.style));
		pPr += wrapRunProperties(
			runPropertiesXml(block.markProps, conversion.charBase(block.style, undefined), conversion, {
				style: false,
			}),
		);
		if (block.endsSection !== undefined) pPr += this.sectionProperties(block.endsSection);
		const inlines = inlinesXml(this, block.inlines, block.style);
		return `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${inlines}</w:p>`;
	}

	private paragraphStyleId(block: DocParagraphBlock): string | undefined {
		const style = this.conversion.styles.style(block.style);
		return style && style.sti !== 0 ? this.conversion.styles.idOf(block.style) : undefined;
	}

	table(block: DocTableBlock): string {
		return tableXml(block, (cell) => {
			const content = this.blocks(cell.blocks);
			return cell.blocks.at(-1)?.type === 'paragraph' ? content : `${content}<w:p/>`;
		});
	}

	blocks(blocks: readonly DocBlock[]): string {
		return blocks
			.map((block) => (block.type === 'paragraph' ? this.paragraph(block) : this.table(block)))
			.join('');
	}

	/** A whole story; never empty, since every container needs a closing paragraph. */
	storyXml(story: DocModelStory): string {
		if (this.textBoxDepth >= MAX_NESTED_TEXT_BOXES) {
			this.conversion.warn('Text boxes nested more than two levels deep were left out.');
			return '<w:p/>';
		}
		this.textBoxDepth++;
		try {
			const content = this.blocks(story.blocks);
			return story.blocks.at(-1)?.type === 'paragraph' ? content : `${content}<w:p/>`;
		} finally {
			this.textBoxDepth--;
		}
	}
}
