/**
 * segments-text-edit.ts: carry one contiguous text edit (typing, a deletion,
 * a paste) onto rich segments, so every character that was not edited keeps
 * exactly the run it was in and the inserted characters take the format of
 * the character before them, as PowerPoint does when typing.
 *
 * Unlike a length-proportional remap this finds the edit itself (the common
 * prefix and suffix of the old and new text, bounded by the caret when the
 * caller knows it), so replacing one word never shifts the run boundaries
 * around it.
 *
 * `\n` in the text is a paragraph break; an existing line break segment
 * (`isLineBreak`) also spells `\n` and is kept as one when it survives.
 *
 * @module render/segments-text-edit
 */
import type { TextSegment, TextStyle } from 'ooxml-core/pptx';

interface Char {
	ch: string;
	/** Index of the source segment, or -1 for an inserted character. */
	seg: number;
	/** Source paragraph index. */
	para: number;
}

/** The text segments spell: paragraph and line breaks are `\n`. */
export function segmentsPlainText(segments: readonly TextSegment[]): string {
	return segments
		.map((segment) => (segment.isParagraphBreak || segment.isLineBreak ? '\n' : segment.text))
		.join('');
}

function explode(segments: readonly TextSegment[]): Char[] {
	const chars: Char[] = [];
	let para = 0;
	segments.forEach((segment, seg) => {
		if (segment.isParagraphBreak) {
			chars.push({ ch: '\n', seg, para });
			para++;
			return;
		}
		if (segment.isLineBreak) {
			chars.push({ ch: '\n', seg, para });
			return;
		}
		for (const ch of segment.text) {
			chars.push({ ch, seg, para });
		}
	});
	return chars;
}

/** A paragraph's own layout and source markup, carried by its first segment. */
interface Layout {
	properties?: TextStyle;
	paragraph?: TextSegment['tableParagraph'];
}

function paragraphLayouts(segments: readonly TextSegment[]): Array<Layout | undefined> {
	const layouts: Array<Layout | undefined> = [];
	let atStart = true;
	for (const segment of segments) {
		if (segment.isParagraphBreak) {
			if (atStart) {
				layouts.push(undefined);
			}
			atStart = true;
			continue;
		}
		if (atStart) {
			layouts.push({
				properties: segment.paragraphProperties,
				paragraph: segment.tableParagraph,
			});
			atStart = false;
		}
	}
	if (atStart) {
		layouts.push(undefined);
	}
	return layouts;
}

/**
 * Apply the edit that turns `segments`' text into `text`. `caret`, when
 * given, is the caret offset in `text` after the edit: the edited range ends
 * there, which settles where a repeated character was typed.
 */
export function editSegmentsText(
	segments: readonly TextSegment[],
	text: string,
	caret?: number,
): TextSegment[] {
	const old = explode(segments);
	const next = Array.from(text);
	if (old.map((c) => c.ch).join('') === text) {
		return segments.map((segment) => ({ ...segment }));
	}
	const maxPrefix = Math.min(old.length, next.length, caret ?? Number.POSITIVE_INFINITY);
	let prefix = 0;
	while (prefix < maxPrefix && old[prefix]!.ch === next[prefix]) {
		prefix++;
	}
	const maxSuffix = Math.min(
		old.length - prefix,
		next.length - prefix,
		caret === undefined ? Number.POSITIVE_INFINITY : Math.max(0, next.length - caret),
	);
	let suffix = 0;
	while (
		suffix < maxSuffix &&
		old[old.length - 1 - suffix]!.ch === next[next.length - 1 - suffix]
	) {
		suffix++;
	}
	// Typed text takes the run before it; at the very start, the run after it.
	const before = old[prefix - 1];
	const after = old[old.length - suffix];
	const host =
		before && !segments[before.seg]?.isParagraphBreak
			? before
			: (after ?? before ?? { ch: '', seg: 0, para: 0 });
	const insertedPara = before
		? segments[before.seg]?.isParagraphBreak
			? before.para + 1
			: before.para
		: 0;
	// Typed text joins its host run, unless that run is atomic (a field or an
	// equation), whose own text must not grow; a typed break is never the host's.
	const hostSegment = segments[host.seg];
	const atomic = hostSegment?.fieldType !== undefined || hostSegment?.equationXml !== undefined;
	const inserted: Char[] = next.slice(prefix, next.length - suffix).map((ch) => ({
		ch,
		seg: atomic || ch === '\n' ? -1 - host.seg : host.seg,
		para: insertedPara,
	}));
	const chars = [...old.slice(0, prefix), ...inserted, ...old.slice(old.length - suffix)];
	return build(segments, chars, paragraphLayouts(segments));
}

/** A segment without its paragraph layout (which only a paragraph's first segment carries). */
function withoutLayout(segment: TextSegment): TextSegment {
	const copy = { ...segment };
	delete copy.paragraphProperties;
	delete copy.tableParagraph;
	return copy;
}

function styleOf(segments: readonly TextSegment[], seg: number): TextStyle {
	const source = segments[seg < 0 ? -1 - seg : seg];
	return { ...(source?.style ?? {}) };
}

function build(
	segments: readonly TextSegment[],
	chars: readonly Char[],
	layouts: ReadonlyArray<Layout | undefined>,
): TextSegment[] {
	const out: TextSegment[] = [];
	let paragraphStart = true;
	let paragraphPara = chars[0]?.para ?? 0;
	let lastSeg = chars[0]?.seg ?? 0;
	const open = (segment: TextSegment, para: number): void => {
		if (paragraphStart) {
			const layout = layouts[para];
			if (layout?.properties) {
				segment.paragraphProperties = { ...layout.properties };
			}
			if (layout?.paragraph) {
				segment.tableParagraph = layout.paragraph;
			}
			paragraphStart = false;
		}
		out.push(segment);
	};
	let run: { segment: TextSegment; seg: number } | undefined;
	for (const char of chars) {
		const source = char.seg >= 0 ? segments[char.seg] : undefined;
		if (char.ch === '\n') {
			run = undefined;
			if (source?.isLineBreak) {
				open({ ...withoutLayout(source), style: { ...source.style } }, char.para);
				continue;
			}
			if (paragraphStart) {
				open({ text: '', style: styleOf(segments, lastSeg) }, paragraphPara);
			}
			out.push({
				...(source?.isParagraphBreak ? withoutLayout(source) : {}),
				text: '\n',
				isParagraphBreak: true,
				style: styleOf(segments, char.seg),
			});
			paragraphStart = true;
			// An inserted break starts a paragraph laid out like the one it split.
			paragraphPara = char.seg >= 0 ? char.para + 1 : char.para;
			lastSeg = char.seg;
			continue;
		}
		lastSeg = char.seg;
		if (run && run.seg === char.seg) {
			run.segment.text += char.ch;
			continue;
		}
		const base = source && !source.isParagraphBreak && !source.isLineBreak ? source : undefined;
		const segment: TextSegment = base
			? { ...withoutLayout(base), text: char.ch, style: { ...base.style } }
			: { text: char.ch, style: styleOf(segments, char.seg) };
		if (char.seg < 0) {
			// Typed text is plain text in the host run's format: never a field,
			// but it keeps that run's own properties (written as an `a:r`).
			delete segment.fieldType;
			const host = segments[-1 - char.seg]?.tableRun;
			if (host && host.tag !== 'a:br') {
				segment.tableRun = host;
			}
		}
		run = { segment, seg: char.seg };
		open(segment, paragraphStart ? paragraphPara : char.para);
	}
	if (paragraphStart) {
		open({ text: '', style: styleOf(segments, lastSeg) }, paragraphPara);
	}
	return out;
}
