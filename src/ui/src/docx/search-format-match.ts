// ProseMirror-level matching and patching for Word Find & Replace formatting criteria
// (`FindFormatCriteria` and `ReplaceFormatPatch` from `ooxml-core/docx`).
import type { Mark, Node as ProseMirrorNode, Schema } from 'prosemirror-model';
import {
	normalizeFormatColor,
	type FindFormatCriteria,
	type ReplaceFormatPatch,
} from 'ooxml-core/docx';
import { translate, type EditorLocale } from './localization';

/** Checks whether the character (or paragraph) at ProseMirror `pos` satisfies `criteria`. */
export function positionMatchesFormat(
	doc: ProseMirrorNode,
	pos: number,
	criteria: FindFormatCriteria,
): boolean {
	if (pos < 0 || pos >= doc.content.size) return false;
	const $pos = doc.resolve(pos);
	const para = $pos.parent;
	if (criteria.alignment !== undefined) {
		const align =
			(para.attrs.align as string | null) ?? (para.attrs.direction === 'rtl' ? 'right' : 'left');
		if (align !== criteria.alignment) return false;
	}
	const marks = $pos.nodeAfter?.isText
		? $pos.nodeAfter.marks
		: $pos.nodeBefore?.isText
			? $pos.nodeBefore.marks
			: $pos.marks();
	if (criteria.styleId !== undefined) {
		const paraStyle = ((para.attrs.style as string | null) || 'Normal').toLowerCase();
		const charStyle = String(
			marks.find((m) => m.type.name === 'characterStyle')?.attrs.id ?? '',
		).toLowerCase();
		const want = criteria.styleId.toLowerCase();
		if (paraStyle !== want && charStyle !== want) return false;
	}
	if (criteria.bold !== undefined && marks.some((m) => m.type.name === 'bold') !== criteria.bold)
		return false;
	if (
		criteria.italic !== undefined &&
		marks.some((m) => m.type.name === 'italic') !== criteria.italic
	)
		return false;
	if (
		criteria.underline !== undefined &&
		marks.some((m) => m.type.name === 'underline') !== criteria.underline
	)
		return false;
	const font = marks.find((m) => m.type.name === 'font');
	if (criteria.fontFamily !== undefined) {
		const actual = String(font?.attrs.family ?? '')
			.trim()
			.toLowerCase();
		if (actual !== criteria.fontFamily.trim().toLowerCase()) return false;
	}
	if (criteria.fontSizePt !== undefined) {
		if (Number(font?.attrs.size) !== criteria.fontSizePt) return false;
	}
	if (criteria.color !== undefined) {
		const actual = font?.attrs.color ? normalizeFormatColor(String(font.attrs.color)) : undefined;
		if (!actual || actual !== normalizeFormatColor(criteria.color)) return false;
	}
	if (criteria.highlight !== undefined) {
		const actual = marks.find((m) => m.type.name === 'highlight')?.attrs.color as
			| string
			| null
			| undefined;
		if (criteria.highlight === 'none') {
			if (actual && actual !== 'none') return false;
		} else if (actual !== criteria.highlight) {
			return false;
		}
	}
	return true;
}

/** Applies run-level fields of `patch` to a ProseMirror mark list. */
export function applyFormatPatchToMarks(
	schema: Schema,
	marks: readonly Mark[],
	patch: ReplaceFormatPatch,
): Mark[] {
	let next = [...marks];
	const toggle = (name: 'bold' | 'italic' | 'underline', value: boolean | undefined) => {
		if (value === undefined) return;
		const type = schema.marks[name];
		if (!type) return;
		next = next.filter((m) => m.type !== type);
		if (value) next.push(type.create());
	};
	toggle('bold', patch.bold);
	toggle('italic', patch.italic);
	toggle('underline', patch.underline);

	if (
		patch.fontFamily !== undefined ||
		patch.fontSizePt !== undefined ||
		patch.color !== undefined
	) {
		const fontType = schema.marks.font;
		if (fontType) {
			const existing = next.find((m) => m.type === fontType);
			const attrs: Record<string, unknown> = {
				family: existing?.attrs.family ?? null,
				size: existing?.attrs.size ?? null,
				color: existing?.attrs.color ?? null,
			};
			if (patch.fontFamily !== undefined) attrs.family = patch.fontFamily;
			if (patch.fontSizePt !== undefined) attrs.size = patch.fontSizePt;
			if (patch.color !== undefined)
				attrs.color = patch.color === null ? null : normalizeFormatColor(patch.color);
			next = next.filter((m) => m.type !== fontType);
			if (attrs.family || attrs.size || attrs.color) next.push(fontType.create(attrs));
		}
	}

	if (patch.highlight !== undefined) {
		const hlType = schema.marks.highlight;
		if (hlType) {
			next = next.filter((m) => m.type !== hlType);
			if (patch.highlight !== null && patch.highlight !== 'none')
				next.push(hlType.create({ color: patch.highlight }));
		}
	}

	if (patch.characterStyleId !== undefined) {
		const csType = schema.marks.characterStyle;
		if (csType) {
			next = next.filter((m) => m.type !== csType);
			if (patch.characterStyleId !== null) next.push(csType.create({ id: patch.characterStyleId }));
		}
	}

	return next;
}

/** Applies paragraph-level fields (`alignment`, `styleId`) of `patch` to paragraph node attrs. */
export function applyFormatPatchToParagraphAttrs(
	attrs: Record<string, unknown>,
	patch: ReplaceFormatPatch,
): Record<string, unknown> {
	if (patch.alignment === undefined && patch.styleId === undefined) return attrs;
	const next: Record<string, unknown> = { ...attrs };
	if (patch.alignment !== undefined) {
		next.align = patch.alignment;
		next.justification = null;
	}
	if (patch.styleId !== undefined) {
		next.style = patch.styleId === null ? '' : patch.styleId;
	}
	return next;
}

/** Concise human-readable summary of active Find or Replace formatting criteria. */
export function formatSummary(
	locale: EditorLocale,
	spec: FindFormatCriteria | ReplaceFormatPatch | undefined,
): string {
	if (!spec) return '';
	const parts: string[] = [];
	if (spec.fontFamily) parts.push(spec.fontFamily);
	if (spec.fontSizePt !== undefined) parts.push(`${spec.fontSizePt} pt`);
	if (spec.bold === true) parts.push(translate(locale, 'Bold'));
	if (spec.bold === false) parts.push(translate(locale, 'Not Bold'));
	if (spec.italic === true) parts.push(translate(locale, 'Italic'));
	if (spec.italic === false) parts.push(translate(locale, 'Not Italic'));
	if (spec.underline === true) parts.push(translate(locale, 'Underline'));
	if (spec.underline === false) parts.push(translate(locale, 'No Underline'));
	if (typeof spec.color === 'string') parts.push(spec.color);
	if (typeof spec.highlight === 'string')
		parts.push(`${translate(locale, 'Highlight')}: ${spec.highlight}`);
	if (spec.alignment) parts.push(spec.alignment);
	if (typeof spec.styleId === 'string') parts.push(spec.styleId);
	return parts.join(', ');
}
