// Lists the MERGEFIELD fields and repeating regions of a document, for the headless merge
// (`mergeFields`): what a service needs to know about a template before it fills it.
import { fieldResultSpans } from '../field-result-spans';
import { fieldTokens } from '../field-instruction';
import type { DocumentModel, Paragraph } from '../model';
import { mapMergeStory } from './cross-paragraph';
import { mergeFieldColumn } from './field-values';
import { listMergeRegions, regionMarker, type MergeRegionInfo } from './regions';
import { forEachBlockList, type HeaderFooterSlot, type MergeStory } from './stories';

export interface MergeFieldInfo {
	/** The column the field reads, as written (`client.2.name`, `First name`). */
	name: string;
	instruction: string;
	story: MergeStory;
	/** In a header or footer, the section slot it fills. */
	slot?: HeaderFooterSlot;
	/** `w:fldSimple`, or a complex field (`w:fldChar`). */
	kind: 'simple' | 'complex';
	/** The switches as written, such as `\* Upper` or `\b "Dear "`. */
	switches: string[];
	/** The region the field repeats in. */
	region?: string;
	/** A locked field (`w:fldLock`) is never filled. */
	locked: boolean;
	/** The field sits in the code of another field (IF, QUOTE...). */
	nested: boolean;
}

function switchesOf(instruction: string): string[] {
	const tokens = fieldTokens(instruction);
	const switches: string[] = [];
	tokens.forEach((token, index) => {
		if (token.kind !== 'switch') return;
		const next = tokens[index + 1];
		const takes = ['*', '@', '#', 'b', 'f'].includes(token.value) && next?.kind === 'text';
		switches.push(
			`\\${token.value}${takes ? ` ${/\s/.test(next.value) ? `"${next.value}"` : next.value}` : ''}`,
		);
	});
	return switches;
}

/** Every MERGEFIELD of `model` (region markers excluded) and every region. */
export function listMergeFields(model: DocumentModel): {
	fields: MergeFieldInfo[];
	regions: MergeRegionInfo[];
} {
	const fields: MergeFieldInfo[] = [];
	// Lists are visited innermost first, a table's cells row by row, so a region in table rows
	// covers the cells between its markers.
	let region: string | undefined;
	forEachBlockList(model, (blocks, list) => {
		const visit = (paragraph: Paragraph) => {
			const spans = fieldResultSpans(paragraph.runs, { resultless: true });
			for (const span of spans) {
				const marker = regionMarker(span.instr);
				if (marker) {
					region = marker.start ? marker.name : undefined;
					continue;
				}
				const name = mergeFieldColumn(span.instr);
				if (!name) continue;
				const nested = spans.some(
					(other) => other !== span && other.from < span.from && other.to > span.to,
				);
				fields.push({
					name,
					instruction: span.instr,
					story: list.story,
					...(list.slot && { slot: list.slot }),
					kind: paragraph.runs[span.from]?.field?.simple ? 'simple' : 'complex',
					switches: switchesOf(span.instr),
					...(region !== undefined && { region }),
					locked: span.locked,
					nested,
				});
			}
			return paragraph;
		};
		// Fields whose code or result crosses paragraphs are read as one joined paragraph.
		const top = blocks.filter((block): block is Paragraph => block.type === 'paragraph');
		mapMergeStory(top, visit);
	});
	return { fields, regions: listMergeRegions(model) };
}
