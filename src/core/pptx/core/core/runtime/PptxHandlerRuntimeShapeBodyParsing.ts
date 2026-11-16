import { XmlObject, TextStyle } from '../../types';
import type { PptxShapeLocks, PptxTextWarpPreset } from '../../types';
import { selectAlternateContentBranch as selectACBranch } from '../../utils/alternate-content';
import {
	autofitFromDrawing,
	bodyFlagsFromDrawing,
	bodyInsetsFromDrawing,
	objectBodyProperties,
} from '../../utils/body-properties-neutral';
import { objectTextElementView } from '../../utils/paragraph-properties-neutral';
import { drawingIntegerValue } from '../../../../drawingml/text-paragraph-codec';
import { parseTextBodyScene3d } from '../../utils/text-body-scene3d';
import { parseTextBodySp3d } from '../../utils/text-body-sp3d';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeBulletParsing';
import type { BodyPropertiesResult } from './PptxHandlerRuntimeTypes';
import { isFiniteNumber } from '../../../../units/finite';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/**
	 * Parse `a:spLocks` attributes into a structured {@link PptxShapeLocks} object.
	 * Returns `undefined` when the node is absent or contains no lock attributes.
	 */
	protected parseShapeLocks(spLocks: XmlObject | undefined): PptxShapeLocks | undefined {
		if (!spLocks) {
			return undefined;
		}

		const locks: PptxShapeLocks = {};
		let hasAny = false;

		const boolAttr = (attr: string): boolean | undefined => {
			const raw = spLocks[attr];
			if (raw === undefined) {
				return undefined;
			}
			const val = String(raw).trim().toLowerCase();
			return val === '1' || val === 'true';
		};

		const set = (attr: string, fn: (v: boolean) => void) => {
			const val = boolAttr(attr);
			if (val !== undefined) {
				fn(val);
				hasAny = true;
			}
		};

		set('@_noGrp', (v) => {
			locks.noGrouping = v;
		});
		set('@_noRot', (v) => {
			locks.noRotation = v;
		});
		set('@_noMove', (v) => {
			locks.noMove = v;
		});
		set('@_noResize', (v) => {
			locks.noResize = v;
		});
		set('@_noTextEdit', (v) => {
			locks.noTextEdit = v;
		});
		set('@_noSelect', (v) => {
			locks.noSelect = v;
		});
		set('@_noChangeAspect', (v) => {
			locks.noChangeAspect = v;
		});
		set('@_noEditPoints', (v) => {
			locks.noEditPoints = v;
		});
		set('@_noAdjustHandles', (v) => {
			locks.noAdjustHandles = v;
		});
		set('@_noChangeArrowheads', (v) => {
			locks.noChangeArrowheads = v;
		});
		set('@_noChangeShapeType', (v) => {
			locks.noChangeShapeType = v;
		});
		// `a:picLocks` only: pictures carry noCrop on top of the shared set.
		set('@_noCrop', (v) => {
			locks.noCrop = v;
		});

		return hasAny ? locks : undefined;
	}

	/**
	 * Parse the `@txBox` attribute from a `p:cNvSpPr` node. Returns `true` /
	 * `false` when the attribute is present, or `undefined` when absent so
	 * callers can distinguish "not a text box" from "unspecified".
	 */
	protected parseTxBoxFlag(cNvSpPr: XmlObject | undefined): boolean | undefined {
		const raw = cNvSpPr?.['@_txBox'];
		if (raw === undefined) {
			return undefined;
		}
		const val = String(raw).trim().toLowerCase();
		return val === '1' || val === 'true';
	}

	/**
	 * Extract body-level text properties from `a:bodyPr` and apply them to the
	 * provided {@link TextStyle}.  Returns linked-textbox info when present.
	 */
	protected applyBodyProperties(
		bodyPr: XmlObject | undefined,
		txBody: XmlObject | undefined,
		textStyle: TextStyle,
	): BodyPropertiesResult {
		const result: BodyPropertiesResult = {};

		// Linked text box chain (a:bodyPr > a:linkedTxbx or direct a:linkedTxbx sibling)
		const linkedTxbx = (bodyPr?.['a:linkedTxbx'] ?? txBody?.['a:linkedTxbx']) as
			| XmlObject
			| undefined;
		const linkedView = objectTextElementView(linkedTxbx);
		if (linkedTxbx && linkedView) {
			const ltxId = drawingIntegerValue(linkedView.attribute('id'));
			if (ltxId !== undefined) {
				result.linkedTxbxId = ltxId;
				result.linkedTxbxSeq = drawingIntegerValue(linkedView.attribute('seq')) ?? 0;
			}
		}

		if (!bodyPr) {
			return result;
		}

		// The attributes and the autofit choice are read by the shared DrawingML
		// body codec (`body-properties-neutral`); text warp, 3D and linked text
		// boxes are read here and by their own readers.
		const body = objectBodyProperties(bodyPr);

		const verticalAlign = this.textVerticalAlignFromDrawingValue(body.anchor);
		if (verticalAlign) {
			textStyle.vAlign = verticalAlign;
		}

		const textDirection = this.textDirectionFromDrawingValue(body.vertical);
		if (textDirection) {
			textStyle.textDirection = textDirection;
		}

		const bodyColumnCount = this.normalizeTextColumnCount(body.columnCount);
		if (bodyColumnCount && bodyColumnCount > 1) {
			textStyle.columnCount = bodyColumnCount;
		}

		if (body.columnSpacingEmu !== undefined && body.columnSpacingEmu > 0) {
			textStyle.columnSpacing = body.columnSpacingEmu / PptxHandlerRuntime.EMU_PER_PX;
		}

		// `@hOverflow` is not a CT_TextBodyProperties attribute, but some
		// producers write it for `@horzOverflow`; it is read as a fallback.
		const hOverflow =
			bodyPr['@_horzOverflow'] !== undefined
				? body.horizontalOverflow
				: String(bodyPr['@_hOverflow'] ?? '').trim();
		if (hOverflow === 'overflow' || hOverflow === 'clip') {
			textStyle.hOverflow = hOverflow;
		}
		const vertOverflow = body.verticalOverflow;
		if (vertOverflow === 'overflow' || vertOverflow === 'clip' || vertOverflow === 'ellipsis') {
			textStyle.vertOverflow = vertOverflow;
		}

		// Text warp preset from a:prstTxWarp
		// The warp preset and its guides come from the shared body codec.
		const warp = body.textWarp;
		if (warp) {
			if (warp.preset !== 'textNoShape') {
				textStyle.textWarpPreset = warp.preset as PptxTextWarpPreset;
			}
			for (const guide of warp.guides ?? []) {
				// Formula is typically "val <number>"
				const valMatch = guide.formula.trim().match(/^val\s+(-?\d+)$/);
				if (!valMatch) {
					continue;
				}
				const val = parseInt(valMatch[1] ?? '', 10);
				if (!isFiniteNumber(val)) {
					continue;
				}
				// The wave and pour presets name their first guide `adj1`, the others `adj`.
				if (guide.name === 'adj' || guide.name === 'adj1') {
					textStyle.textWarpAdj = val;
				} else if (guide.name === 'adj2') {
					textStyle.textWarpAdj2 = val;
				}
			}
		}

		// 3D text body properties (a:bodyPr/a:sp3d)
		parseTextBodySp3d(bodyPr, textStyle, (colorNode) => this.parseColor(colorNode));

		parseTextBodyScene3d(bodyPr, textStyle);

		// Auto-fit
		Object.assign(textStyle, autofitFromDrawing(body.autofit));

		// Text wrapping mode from a:bodyPr/@wrap
		const wrapAttr = body.wrap?.toLowerCase();
		if (wrapAttr === 'none') {
			textStyle.textWrap = 'none';
		} else if (wrapAttr === 'square') {
			textStyle.textWrap = 'square';
		}

		// Body text insets (lIns, tIns, rIns, bIns; EMU values) and the flags
		Object.assign(textStyle, bodyInsetsFromDrawing(body), bodyFlagsFromDrawing(body));

		return result;
	}

	/**
	 * Select the appropriate branch from an mc:AlternateContent element.
	 * Delegates to the standalone alternate-content utility which checks
	 * mc:Choice/@_Requires against the full set of supported namespaces
	 * and handles nested AlternateContent blocks.
	 */
	protected selectAlternateContentBranch(ac: XmlObject): XmlObject | undefined {
		return selectACBranch(ac);
	}
}
