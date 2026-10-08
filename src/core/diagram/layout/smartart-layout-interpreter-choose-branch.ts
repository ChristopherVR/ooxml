/**
 * SmartArt DiagramML interpreter - `dgm:choose` branch selection (top-level
 * and nested).
 *
 * Split out of `smartart-layout-interpreter-choose-algorithm.ts` (the
 * file-size budget): the two functions that pick WHICH branch is active, for
 * a layoutNode's own parsed `dgm:choose` ({@link activeBranch}, returning the
 * branch's raw-XML slot as the model holds it) and for one living entirely
 * inside an already-active branch's body ({@link nestedChooseBranch}, over
 * the ordered-XML tree) - see that module's own doc comment for how
 * `branchAlg` uses both.
 */

import type { DiagramChoose } from '../model';
import type { OrderedXmlElement } from './smartart-choose-xml';
import { firstChildNamed, groupedChildren } from './smartart-choose-xml';
import { parseOrderedWhen } from './smartart-layout-control-flow';
import { evaluateWhen } from './smartart-layout-interpreter-when';
import type { WhenContext } from './smartart-layout-interpreter-when';

/**
 * Resolve the raw XML of the active `dgm:choose` branch for a node count, or
 * `undefined` when the choose is not decidable (an earlier branch is
 * undecidable) or no branch applies. DiagramML picks the first matching `if` in
 * order, so an undecidable earlier branch forces a bail.
 */
export function activeBranch<R>(
	choose: DiagramChoose<R>,
	nodeCount: number,
	context: WhenContext,
): R | undefined {
	for (const when of choose.when) {
		const result = evaluateWhen(when, nodeCount, context);
		if (result === undefined) {
			return undefined;
		}
		if (result) {
			return when.rawXml;
		}
	}
	return choose.otherwise?.rawXml ?? undefined;
}

/**
 * Resolve a NESTED `dgm:choose` (one living entirely inside an already-active
 * branch's raw XML, e.g. `basic-radial--hier5.pptx`'s `stAng` choose nested
 * inside its `dir="norm"` branch) to its own winning branch, or `undefined`
 * when it is not decidable - never a raw `dgm:if`/`dgm:else` `branchAlg`'s
 * blind walk would otherwise recurse into UNCONDITIONALLY. Only a
 * layoutNode's OWN direct `dgm:choose` children are parsed into the typed
 * model at load time; one nested inside a branch is only ever reachable
 * here, parsed on the fly with the SAME `dgm:if` attribute parser the
 * top-level reader uses, so it decides with the exact same rules (including
 * compound `@axis` navigation - see `smartart-layout-interpreter-when.ts`).
 */
export function nestedChooseBranch(
	choose: OrderedXmlElement,
	nodeCount: number,
	context: WhenContext,
): OrderedXmlElement | undefined {
	for (const child of groupedChildren(choose)) {
		if (child.name !== 'if') {
			continue;
		}
		const when = parseOrderedWhen(child);
		if (!when) {
			continue;
		}
		const result = evaluateWhen(when, nodeCount, context);
		if (result === undefined) {
			// An earlier `if` is undecidable: DiagramML picks the first
			// matching branch in document order, so an undecidable one
			// forces the WHOLE nested choose undecidable too - matching
			// `activeBranch`'s own top-level rule.
			return undefined;
		}
		if (result) {
			return child;
		}
	}
	return firstChildNamed(choose, 'else');
}
