import { XmlObject } from '../../types';
import type { PptxElementWithText, TextStyle } from '../../types';
import type {
	DrawingAutofit,
	DrawingBodyProperties,
	DrawingTextWarpGuide,
} from '../../../../drawingml/text-model';
import {
	assignBodyAttributes,
	objectAutofit,
	objectTextWarp,
} from '../../utils/body-properties-neutral';
import { writeBodyPrBooleanAttrs } from '../../utils/body-properties-parser';
import { applyTextBodyScene3d } from '../../utils/text-body-scene3d';
import { applyTextBodySp3d } from '../../utils/text-body-sp3d-writer';
import {
	elementBodyPropertyEditKeys,
	shouldWriteBodyProperty,
	type ElementBodyPropertyKey,
	styleForBodyPropertyWrite,
} from './element-body-properties';
import {
	applyElementParagraphGeometryToListStyle,
	elementParagraphGeometryEdits,
	hasElementParagraphGeometry,
	withoutElementParagraphGeometry,
} from './element-paragraph-geometry';
import { preserveParagraphScopedState } from './paragraph-scoped-segment-state';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveEffectsWriter';
import { withRunSlideJumpTarget, withSegmentSlideJumpTargets } from './run-slide-jump-targets';
import type { SlideTargetAt } from './run-slide-jump-targets';
import { buildParagraphPropertiesXml } from './PptxHandlerRuntimeSaveParagraphHelpers';

/** The presets whose first guide is `adj1` (the two-guide waves and pours); every other preset names it `adj`. */
const ADJ1_WARP_PRESETS: ReadonlySet<string> = new Set([
	'textWave1',
	'textWave2',
	'textWave4',
	'textDoubleWave1',
	'textArchUpPour',
	'textArchDownPour',
	'textCirclePour',
	'textButtonPour',
]);

function firstWarpGuideName(preset: string | undefined): 'adj' | 'adj1' {
	return preset !== undefined && ADJ1_WARP_PRESETS.has(preset) ? 'adj1' : 'adj';
}

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Build and write the text body (`p:txBody`) for shapes, text boxes,
	 * and connectors that carry text content.
	 *
	 * Handles bodyPr attributes (vAlign, textDirection, columns, spacing,
	 * overflow, autoFit, body insets, text wrap, warp, 3D text, linked
	 * text chains), and delegates paragraph creation to
	 * `createParagraphsFromTextContent`.
	 */
	protected applyTextBodyContent(
		shape: XmlObject,
		el: PptxElementWithText,
		resolveHyperlinkRelationshipId: (target: string) => string | undefined,
		getSlideRelationshipMap: () => Map<string, string>,
		slideTargetAt?: SlideTargetAt,
	): void {
		const hasEditableTextContent =
			typeof el.text === 'string' || (el.textSegments?.length ?? 0) > 0;
		if (!hasEditableTextContent) {
			return;
		}

		// A footer / header / date / slide-number placeholder whose text the loader
		// INHERITED from the slide master must go back out the way PowerPoint
		// writes it: an empty body on the slide, so the string keeps coming from
		// the master and the Header & Footer dialog keeps owning it. Re-emitting
		// the resolved string would pin this slide to the text the master happened
		// to hold at load. An edit changes `el.text`, so the equality check still
		// lets a genuine per-slide override through to the writer below.
		if (
			el.inheritedPlaceholderText !== undefined &&
			el.text === el.inheritedPlaceholderText &&
			el.rawXml !== undefined
		) {
			return;
		}

		if (!shape['p:txBody']) {
			shape['p:txBody'] = {
				'a:bodyPr': {},
				'a:lstStyle': {},
				'a:p': [],
			};
		}
		const txBody = shape['p:txBody'] as XmlObject;
		const bodyPr = (txBody['a:bodyPr'] || {}) as XmlObject;

		// Which `a:bodyPr` scoped fields were actually authored/edited, as
		// opposed to merely resolved by the placeholder/layout/master
		// cascade at load time. `undefined` means there is no baseline to
		// diff against (SDK-built text), so every field below is written in
		// full, exactly as before this diff existed. See
		// element-body-properties.ts for why re-emitting an unedited field
		// pins an inherited value onto the slide.
		const bodyEdits = elementBodyPropertyEditKeys(el.textStyle);

		// Vertical anchor
		if (shouldWriteBodyProperty(bodyEdits, 'vAlign')) {
			const verticalAnchor = this.textVerticalAlignToDrawingValue(el.textStyle?.vAlign);
			if (verticalAnchor) {
				assignBodyAttributes(bodyPr, { anchor: verticalAnchor });
			} else {
				delete bodyPr['@_anchor'];
			}
		}

		// Text direction
		if (shouldWriteBodyProperty(bodyEdits, 'textDirection')) {
			const bodyTextDirection = this.textDirectionToDrawingValue(el.textStyle?.textDirection);
			if (bodyTextDirection) {
				assignBodyAttributes(bodyPr, { vertical: bodyTextDirection });
			} else {
				delete bodyPr['@_vert'];
			}
		}

		// Column count
		if (shouldWriteBodyProperty(bodyEdits, 'columnCount')) {
			const bodyColumnCount = this.normalizeTextColumnCount(el.textStyle?.columnCount);
			if (bodyColumnCount && bodyColumnCount > 1) {
				assignBodyAttributes(bodyPr, { columnCount: bodyColumnCount });
			} else {
				delete bodyPr['@_numCol'];
			}
		}

		// Column spacing
		if (
			shouldWriteBodyProperty(bodyEdits, 'columnSpacing') &&
			el.textStyle?.columnSpacing !== undefined
		) {
			assignBodyAttributes(bodyPr, {
				columnSpacingEmu: el.textStyle.columnSpacing * PptxHandlerRuntime.EMU_PER_PX,
			});
		}

		// Overflow
		if (shouldWriteBodyProperty(bodyEdits, 'hOverflow') && el.textStyle?.hOverflow) {
			assignBodyAttributes(bodyPr, { horizontalOverflow: el.textStyle.hOverflow });
			delete bodyPr['@_hOverflow'];
		}
		if (shouldWriteBodyProperty(bodyEdits, 'vertOverflow') && el.textStyle?.vertOverflow) {
			assignBodyAttributes(bodyPr, { verticalOverflow: el.textStyle.vertOverflow });
		}

		// Auto-fit / shrink-to-fit — a mutually-exclusive node choice, so the
		// whole decision is skipped unless one of its driving fields changed;
		// touching it based on an inherited `autoFit`/`autoFitMode` is what
		// turned an inherited `a:normAutofit` into `a:spAutoFit`.
		if (
			bodyEdits === undefined ||
			(['autoFitMode', 'autoFit', 'autoFitFontScale', 'autoFitLineSpacingReduction'] as const).some(
				(key) => bodyEdits.has(key),
			)
		) {
			this.applyAutoFitToBodyPr(bodyPr, el);
		}

		// Body text insets
		this.applyBodyInsets(bodyPr, el, bodyEdits);

		// Text wrapping mode
		if (shouldWriteBodyProperty(bodyEdits, 'textWrap')) {
			const wrap = el.textStyle?.textWrap;
			if (wrap === 'none' || wrap === 'square') {
				assignBodyAttributes(bodyPr, { wrap });
			}
		}

		// Additional bodyPr boolean attributes. `writeBodyPrBooleanAttrs` only
		// ever SETS an attribute when its field is defined (never deletes), so
		// filtering the unedited fields out of the style handed to it is
		// enough to leave those attributes untouched.
		writeBodyPrBooleanAttrs(bodyPr, styleForBodyPropertyWrite(el.textStyle, bodyEdits));

		// Text warp preset
		if (el.textStyle?.textWarpPreset) {
			// Round-trip adjustment values (adj / adj2). PowerPoint always writes
			// `a:avLst`, empty when the warp keeps its default adjustments;
			// omitting it lost markup on every rewrite.
			const guides: DrawingTextWarpGuide[] = [];
			if (el.textStyle.textWarpAdj !== undefined && Number.isFinite(el.textStyle.textWarpAdj)) {
				guides.push({
					name: firstWarpGuideName(el.textStyle.textWarpPreset),
					formula: `val ${el.textStyle.textWarpAdj}`,
				});
			}
			if (el.textStyle.textWarpAdj2 !== undefined && Number.isFinite(el.textStyle.textWarpAdj2)) {
				guides.push({ name: 'adj2', formula: `val ${el.textStyle.textWarpAdj2}` });
			}
			bodyPr['a:prstTxWarp'] = objectTextWarp({ preset: el.textStyle.textWarpPreset, guides });
		} else {
			delete bodyPr['a:prstTxWarp'];
		}

		applyTextBodyScene3d(bodyPr, el.textStyle);
		// 3D text body
		applyTextBodySp3d(bodyPr, el.textStyle);

		// Linked text box chain round-trip
		if (el.linkedTxbxId !== undefined && Number.isFinite(el.linkedTxbxId)) {
			bodyPr['a:linkedTxbx'] = {
				'@_id': String(el.linkedTxbxId),
				'@_seq': String(el.linkedTxbxSeq ?? 0),
			};
		} else {
			delete bodyPr['a:linkedTxbx'];
		}

		// Resolve text value and segments
		const textValueForSave = this.getTextValueForSave(el.text, el.textSegments);
		let textSegmentsForSave = el.textSegments;
		if (typeof el.text === 'string' && this.areTextSegmentsUniform(el.textSegments)) {
			textSegmentsForSave = undefined;
			const existingTextSegments = this.extractTextSegmentsFromTxBodyForRewrite(
				txBody,
				el.textStyle,
				getSlideRelationshipMap(),
			);
			if (existingTextSegments.length > 1 && this.hasMixedTextStyles(existingTextSegments)) {
				textSegmentsForSave = this.remapEditedTextToExistingStyles(
					existingTextSegments,
					textValueForSave,
					el.textStyle,
				);
			}
			// "Uniform" is a RUN-scope judgement (every `a:rPr` would come out
			// the same), so collapsing the runs is legitimate. The paragraph
			// scope is not its to discard: `a:pPr` geometry, the outline level,
			// the bullet and `a:endParaRPr` all ride on the first segment of
			// each paragraph, and the segmentless path supplies none of them, so
			// dropping the list flattened every authored `a:pPr` to `<a:pPr/>`.
			textSegmentsForSave = preserveParagraphScopedState(
				textSegmentsForSave,
				textValueForSave,
				el.textSegments,
			);
		}

		txBody['a:bodyPr'] = bodyPr;
		// A run slide jump made in code names its slide by index only; point it
		// at the slide part so the run writer adds the relationship.
		const onMissing = (index: number) => this.reportMissingSlideJumpTarget(el.id, index);
		txBody['a:p'] = this.createParagraphsFromTextContent(
			textValueForSave,
			withRunSlideJumpTarget(
				this.routeElementParagraphGeometry(txBody, el.textStyle),
				slideTargetAt,
				onMissing,
			),
			withSegmentSlideJumpTargets(textSegmentsForSave, slideTargetAt, onMissing),
			resolveHyperlinkRelationshipId,
		);
	}

	/**
	 * Move the element-scope paragraph geometry out of every `a:pPr` and into
	 * the text body's `a:lstStyle > a:lvl1pPr`, returning the element style with
	 * that geometry removed so the paragraph writer no longer broadcasts it.
	 *
	 * See `element-paragraph-geometry.ts` for why the `a:lstStyle` slot is the
	 * one that makes the "authored versus resolved" question answerable at all.
	 */
	private routeElementParagraphGeometry(
		txBody: XmlObject,
		textStyle: TextStyle | undefined,
	): TextStyle | undefined {
		const edits = elementParagraphGeometryEdits(textStyle);
		if (edits === undefined) {
			// No load-time snapshot: this text was not parsed from a deck, so the
			// element style is its only description and the paragraph writer must
			// keep writing it out. Nothing changes for SDK-built decks.
			return textStyle;
		}
		if (hasElementParagraphGeometry(edits)) {
			applyElementParagraphGeometryToListStyle(
				txBody,
				buildParagraphPropertiesXml(edits, this.textAlignToDrawingValue(edits.align), undefined, {
					spacingBefore: this.createParagraphSpacingXmlFromPx(edits.paragraphSpacingBefore),
					spacingAfter: this.createParagraphSpacingXmlFromPx(edits.paragraphSpacingAfter),
					lineSpacing: this.createLineSpacingXmlFromMultiplier(edits.lineSpacing),
					lineSpacingExactPt: edits.lineSpacingExactPt,
				}),
			);
		}
		// Everything else the element style says about paragraphs came out of the
		// cascade, so it is dropped rather than stamped onto every `a:pPr`: the
		// paragraphs that authored geometry still emit their own, and the rest go
		// back to inheriting it.
		return withoutElementParagraphGeometry(textStyle);
	}

	/** Apply auto-fit mode settings to bodyPr (node values from the shared body codec). */
	private applyAutoFitToBodyPr(bodyPr: XmlObject, el: PptxElementWithText): void {
		const style = el.textStyle;
		const normal = (fontScale: number | undefined): DrawingAutofit => ({
			kind: 'normal',
			...(fontScale !== undefined && fontScale < 1 ? { fontScale } : {}),
			...(style?.autoFitLineSpacingReduction !== undefined && style.autoFitLineSpacingReduction > 0
				? { lineSpacingReduction: style.autoFitLineSpacingReduction }
				: {}),
		});
		const put = (autofit: DrawingAutofit) => {
			const { key, node } = objectAutofit(autofit);
			bodyPr[key] = node;
		};
		if (style?.autoFitMode !== undefined) {
			delete bodyPr['a:spAutoFit'];
			delete bodyPr['a:normAutofit'];
			delete bodyPr['a:noAutofit'];

			if (style.autoFitMode === 'shrink') {
				put({ kind: 'shape' });
			} else if (style.autoFitMode === 'normal') {
				put(normal(style.autoFitFontScale));
			} else if (style.autoFitMode === 'none') {
				put({ kind: 'none' });
			}
		} else if (style?.autoFit) {
			// Legacy path: keep backward compat
			if (!bodyPr['a:spAutoFit'] && !bodyPr['a:normAutofit']) {
				if (style.autoFitFontScale !== undefined && style.autoFitFontScale < 1) {
					put(normal(style.autoFitFontScale));
				} else {
					put({ kind: 'shape' });
				}
			}
			if (bodyPr['a:normAutofit'] && style.autoFitFontScale !== undefined) {
				const { node } = objectAutofit({ kind: 'normal', fontScale: style.autoFitFontScale });
				Object.assign(bodyPr['a:normAutofit'] as XmlObject, node);
			}
		} else if (style?.autoFit === false && style?.autoFitMode === undefined) {
			delete bodyPr['a:spAutoFit'];
			delete bodyPr['a:normAutofit'];
		}
	}

	/** Apply body insets (margin) to bodyPr. */
	private applyBodyInsets(
		bodyPr: XmlObject,
		el: PptxElementWithText,
		bodyEdits: ReturnType<typeof elementBodyPropertyEditKeys>,
	): void {
		const insets: DrawingBodyProperties = {};
		const inset = (key: ElementBodyPropertyKey, value: number | undefined): number | undefined =>
			shouldWriteBodyProperty(bodyEdits, key) && typeof value === 'number' && Number.isFinite(value)
				? value * PptxHandlerRuntime.EMU_PER_PX
				: undefined;
		const left = inset('bodyInsetLeft', el.textStyle?.bodyInsetLeft);
		if (left !== undefined) insets.insetLeftEmu = left;
		const top = inset('bodyInsetTop', el.textStyle?.bodyInsetTop);
		if (top !== undefined) insets.insetTopEmu = top;
		const right = inset('bodyInsetRight', el.textStyle?.bodyInsetRight);
		if (right !== undefined) insets.insetRightEmu = right;
		const bottom = inset('bodyInsetBottom', el.textStyle?.bodyInsetBottom);
		if (bottom !== undefined) insets.insetBottomEmu = bottom;
		assignBodyAttributes(bodyPr, insets);
	}
}
