/**
 * SmartArt DiagramML interpreter - `dgm:choose`-wrapped `dgm:alg` resolution.
 *
 * Split out of `smartart-layout-interpreter-flow.ts` to keep that file under
 * the repo's per-file line budget: this half resolves a decidable `dgm:choose`
 * to its winning branch's arrangement algorithm - the TYPE alone
 * ({@link chooseAlgType}, the pre-existing entry point) or the type PLUS every
 * `dgm:param` it declares ({@link chooseAlgorithm}).
 *
 * The full-algorithm resolution matters because the typed layout model only
 * parses a DIRECT `dgm:alg` child of a layoutNode into that node's own
 * `algorithm` field; a `dgm:alg` living entirely inside a `dgm:choose`
 * (mirroring `dir="rtl"`, the common real-world shape) leaves that field
 * `undefined` at parse time. `discoverArrangement`
 * (`smartart-layout-interpreter-model.ts`) uses {@link chooseAlgorithm} to
 * populate the chosen arranger node's `algorithm` with the WINNING branch's
 * real params, not just its type, when it falls back to the choose-wrapped
 * node itself.
 *
 * Every entry point takes a {@link RawXmlView} that turns the model's raw
 * branch slot into the ordered-XML element the walk reads (see
 * `smartart-choose-xml.ts`).
 */

import type { DiagramLayoutAlgorithm, DiagramLayoutNode } from '../model';
import type { OrderedXmlElement, RawXmlView } from './smartart-choose-xml';
import { algorithmParameters, groupedChildren } from './smartart-choose-xml';
import { activeBranch, nestedChooseBranch } from './smartart-layout-interpreter-choose-branch';
import { boundedCompositeAlg } from './smartart-layout-interpreter-choose-composite';
import type { WhenContext } from './smartart-layout-interpreter-when';

/**
 * Structural algorithm types a decidable choose branch may select. Includes
 * `hierChild`/`hierRoot`: a genuine org-chart layoutDef (ECMA-376 orgChart1)
 * wraps its OWN root hierarchy algorithm in a `dgm:choose` picking between
 * `linDir`/`hierBranch` variants, not a bare `dgm:alg` - excluding them here
 * meant a choose-wrapped hierarchy was never recognised at all (see
 * `smartart-layout-interpreter-model.ts`'s `discoverArrangement`, which
 * special-cases these two types to set `hierarchy` rather than `chosen`).
 *
 * Deliberately EXCLUDES `composite`: recognising it here would let
 * `discoverArrangement`'s `else if (kind === 'composite' ...)` branch reach
 * a choose-wrapped `composite` algorithm (`Basic Venn`/`Interconnected
 * Rings`/`Theme Picture Accent/Grid/Alternating Accent`/`Bubble Picture
 * List` all wrap theirs this way) - a genuine, still-open fix - but TRIED
 * and reverted: `branchAlg`'s blind recursive search (see its own doc
 * comment) also then matches an UNRELATED nested `composite` reached
 * through the SAME walk (a per-item template, or a different node's own
 * choose in a hub+satellite `cycle` family via `detectHubExpansion` -
 * `smartart-layout-interpreter-hub.ts`), stealing the decision away from a
 * genuinely earlier `cycle`/`lin` alg before it is ever reached - measured
 * regression on `Upward Arrow`/`Phased Process`/`Circle Relationship`/
 * `Opposing Ideas`/`Stacked Venn`/`Radial Picture List` even after bounding
 * the search at a `dgm:layoutNode` boundary (which ALSO regressed
 * `Continuous Cycle`/`Segmented Process`/`Accented Picture`/`Nested Target`
 * on its own, changing which alg a `cycle`/`lin`/`snake`/`pyra` branch finds
 * too). `boundedCompositeAlg` is the narrower fallback that was kept.
 */
const CHOOSE_ALG_TYPES = new Set(['lin', 'cycle', 'pyra', 'snake', 'hierChild', 'hierRoot']);

/** A recognised structural `dgm:alg` found inside a branch's XML, its type plus the element (for its `dgm:param`s). */
interface FoundBranchAlg {
	type: string;
	raw: OrderedXmlElement;
}

/**
 * First recognised structural `dgm:alg` declared inside a branch's XML,
 * carrying its element so a caller can also read its `dgm:param`s (see
 * {@link chooseAlgorithm}) - a type-only search generalised so a
 * choose-wrapped `dgm:alg`'s params (`grDir`/`flowDir`/`contDir`/`off`/
 * `linDir`/...) are not silently lost. Measured against `basic-block-list--
 * flat3.pptx`: its `snake` algorithm lives entirely inside a `dgm:choose`
 * (mirroring `dir="rtl"`), so `discoverArrangement`'s chosen arranger node
 * previously kept its OWN (parse-time, choose-blind) `undefined` `algorithm`
 * field - every `dgm:param` read silently fell back to its generic default
 * instead of the declared value.
 *
 * A NESTED `dgm:choose` inside the branch (see {@link nestedChooseBranch}) is
 * evaluated, not blindly recursed into: COM-verified regression against
 * `basic-radial--hier5.pptx`, whose `dir="norm"` branch nests a SECOND
 * choose picking `stAng` by satellite count - the old blind walk always
 * found the FIRST nested `if`'s `dgm:alg` regardless of whether that
 * branch's own condition was true. An undecidable nested choose is skipped
 * entirely (no fallback to guessing one of its branches).
 */
function branchAlg(
	raw: OrderedXmlElement | undefined,
	nodeCount: number,
	context: WhenContext,
	allowedTypes: ReadonlySet<string> = CHOOSE_ALG_TYPES,
): FoundBranchAlg | undefined {
	if (!raw) {
		return undefined;
	}
	let found: FoundBranchAlg | undefined;
	const visit = (element: OrderedXmlElement): void => {
		for (const child of groupedChildren(element)) {
			if (found !== undefined) {
				return;
			}
			if (child.name === 'alg') {
				const type = child.attrs['type'] ?? '';
				if (allowedTypes.has(type)) {
					found = { type, raw: child };
					return;
				}
			} else if (child.name === 'choose') {
				const winningBranch = nestedChooseBranch(child, nodeCount, context);
				if (winningBranch) {
					visit(winningBranch);
				}
			} else {
				visit(child);
			}
		}
	};
	visit(raw);
	return found;
}

/** The algorithm a found `dgm:alg` element declares, with its `dgm:param`s when it has any. */
function foundAlgorithm(found: FoundBranchAlg): DiagramLayoutAlgorithm {
	const parameters = algorithmParameters(found.raw);
	return { type: found.type, ...(parameters.length > 0 ? { parameters } : {}) };
}

/**
 * Resolve a decidable `dgm:choose` on `node` to the structural algorithm type
 * it selects, or `undefined` when no choose is decidable (in which case the
 * caller keeps the blind first-recognised-alg behaviour). Decidable on
 * `func="cnt"` from `nodeCount` alone, or on `func="var"` when `context`
 * carries `presLayoutVars`; `pos`/`revPos`/`posEven`/`posOdd`/`depth`/
 * `maxDepth` are decidable too when `context` supplies the declaring layout
 * node's own tree location. `context` defaults to `{}` (those functions stay
 * undecidable then).
 */
export function chooseAlgType<R>(
	view: RawXmlView<R>,
	node: DiagramLayoutNode<R>,
	nodeCount: number,
	context: WhenContext = {},
): string | undefined {
	return chooseAlgorithm(view, node, nodeCount, context)?.type;
}

/**
 * Same decidability rules as {@link chooseAlgType}, but resolves the winning
 * branch's FULL algorithm (type plus every `dgm:param`), not just its type -
 * see {@link branchAlg}'s doc comment for why a type-only resolution loses a
 * choose-wrapped arranger's own params entirely. A `composite` found by
 * `boundedCompositeAlg` is the fallback, tried only once every branch's own
 * structural search has found nothing.
 */
export function chooseAlgorithm<R>(
	view: RawXmlView<R>,
	node: DiagramLayoutNode<R>,
	nodeCount: number,
	context: WhenContext = {},
): DiagramLayoutAlgorithm | undefined {
	if (!node.choose || node.choose.length === 0) {
		return undefined;
	}
	for (const choose of node.choose) {
		const found = branchAlg(view(activeBranch(choose, nodeCount, context)), nodeCount, context);
		if (found) {
			return foundAlgorithm(found);
		}
	}
	for (const choose of node.choose) {
		const found = boundedCompositeAlg(
			view(activeBranch(choose, nodeCount, context)),
			nodeCount,
			context,
		);
		if (found) {
			return foundAlgorithm(found);
		}
	}
	return undefined;
}

/**
 * Resolve a decidable `dgm:choose` to a WINNING branch algorithm whose type
 * belongs to `allowedTypes` - a generalisation of {@link chooseAlgType}/
 * {@link chooseAlgorithm} for a caller that needs a type OUTSIDE their own
 * `CHOOSE_ALG_TYPES` whitelist (e.g. `tx`, for hierarchy generation-template
 * detection - `smartart-hierarchy-generation-templates.ts`).
 *
 * Deliberately kept SEPARATE rather than widening `CHOOSE_ALG_TYPES` itself:
 * that whitelist also gates `discoverArrangement`'s own ARRANGEMENT dispatch,
 * and widening it there for an unrelated type has previously regressed
 * unrelated fixtures. No `composite` fallback here.
 */
export function chooseAlgorithmOfType<R>(
	view: RawXmlView<R>,
	node: DiagramLayoutNode<R>,
	nodeCount: number,
	allowedTypes: ReadonlySet<string>,
	context: WhenContext = {},
): DiagramLayoutAlgorithm | undefined {
	if (!node.choose || node.choose.length === 0) {
		return undefined;
	}
	for (const choose of node.choose) {
		const found = branchAlg(
			view(activeBranch(choose, nodeCount, context)),
			nodeCount,
			context,
			allowedTypes,
		);
		if (found) {
			return foundAlgorithm(found);
		}
	}
	return undefined;
}
