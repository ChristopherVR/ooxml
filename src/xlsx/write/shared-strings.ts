import { NS } from '../../xml/index.js';
import type { RichTextRun } from '../model.js';
import { fontXml } from './style-xml.js';
import { XML_HEADER, tElement } from './xml-out.js';

/** The `<si>`/`<is>` body of a plain or rich string. */
export function richStringXml(text: string, runs: readonly RichTextRun[] | undefined): string {
	if (!runs?.length) return tElement(text);
	return runs
		.map(
			(run) =>
				`<r>${run.font && Object.keys(run.font).length ? fontXml(run.font, 'rPr') : ''}${tElement(run.text)}</r>`,
		)
		.join('');
}

/** Collects the shared string table while worksheets are written. */
export class SharedStringTable {
	private readonly items: string[] = [];
	private readonly index = new Map<string, number>();
	private references = 0;

	id(text: string, runs?: readonly RichTextRun[]): number {
		this.references++;
		const body = richStringXml(text, runs);
		let id = this.index.get(body);
		if (id === undefined) {
			id = this.items.length;
			this.items.push(body);
			this.index.set(body, id);
		}
		return id;
	}

	get size(): number {
		return this.items.length;
	}

	xml(): string {
		const items = this.items.map((body) => `<si>${body}</si>`).join('');
		return `${XML_HEADER}<sst xmlns="${NS.x}" count="${this.references}" uniqueCount="${this.items.length}">${items}</sst>`;
	}
}
