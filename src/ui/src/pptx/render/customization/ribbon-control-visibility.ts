/**
 * Group- and control-level ribbon hiding.
 *
 * Every binding tags its ribbon markup with the catalogue ids
 * (`ribbon-control-catalog.ts`): the element that wraps a group carries
 * `data-ribbon-group="<tab>.<group>"` and the element that IS a control (its
 * button, split button, select or gallery trigger) carries
 * `data-ribbon-control="<tab>.<group>.<control>"`. A group whose markup has no
 * wrapper of its own gets a `display: contents` wrapper, so tagging never
 * changes layout.
 *
 * The binding then renders ONE `<style>` element with
 * {@link ribbonCustomizationCss} for its root. Hiding is therefore decided
 * here, once, from the resolved customisation, and the five bindings cannot
 * disagree about which markup a given id removes. Components that prefer to
 * skip rendering can ask {@link isRibbonGroupVisible} /
 * {@link isRibbonControlVisible} instead; both routes read the same sets.
 *
 * @module render/customization/ribbon-control-visibility
 */
import { HOME_DRAWING_ARRANGE_ITEMS } from '../ribbon-drawing-arrange';
import { FORMAT_SECTION_ATTRIBUTE, FORMAT_SECTION_IDS } from '../inspector/format-sections';
import { INSPECTOR_CARD_IDS } from './inspector-customization';
import type { ResolvedCustomization } from './customization-resolve';
import { RIBBON_CONTROL_IDS, RIBBON_GROUP_IDS } from './ribbon-control-ids';
import type { RibbonControlId, RibbonGroupId } from './ribbon-control-ids';

/** Attribute naming a ribbon group on its wrapper element. */
export const RIBBON_GROUP_ATTR = 'data-ribbon-group';
/** Attribute naming a ribbon control on its element. */
export const RIBBON_CONTROL_ATTR = 'data-ribbon-control';
/**
 * Attribute a binding puts on its viewer root so the generated rules only
 * reach its own ribbon when two viewers share a page. The value is any
 * per-instance token.
 */
export const RIBBON_SCOPE_ATTR = 'data-pptx-ribbon-scope';

const GROUP_ID_SET: ReadonlySet<string> = new Set(RIBBON_GROUP_IDS);
const CONTROL_ID_SET: ReadonlySet<string> = new Set(RIBBON_CONTROL_IDS);
const SMART_ART_ARRANGE_GROUP = 'smartArtFormat.arrange';

/** True when `id` names a catalogued ribbon group. */
export function isRibbonGroupId(id: string): id is RibbonGroupId {
	return GROUP_ID_SET.has(id);
}

/** True when `id` names a catalogued ribbon control. */
export function isRibbonControlId(id: string): id is RibbonControlId {
	return CONTROL_ID_SET.has(id);
}

/** True unless the host hid this group. */
export function isRibbonGroupVisible(resolved: ResolvedCustomization, id: RibbonGroupId): boolean {
	return !resolved.hiddenRibbonGroups.has(id);
}

/**
 * Deprecated ids that stand for several controls. `insert.text.field` was the Header, Date and
 * Slide Number menu; hiding it still hides the two buttons that replaced it.
 */
export const RIBBON_CONTROL_ALIASES: Readonly<Record<string, readonly RibbonControlId[]>> = {
	'insert.text.field': ['insert.text.dateTime', 'insert.text.slideNumber'],
	'home.arrange': [
		'home.drawing.arrange',
		...HOME_DRAWING_ARRANGE_ITEMS.map((row) => row.attrs['data-ribbon-control']),
	],
	'home.arrange.align': HOME_DRAWING_ARRANGE_ITEMS.filter(
		({ action }) => action === 'align' || action === 'distribute',
	).map((row) => row.attrs['data-ribbon-control']),
};

/** True unless the host hid this control, an alias that covers it, or the group it sits in. */
export function isRibbonControlVisible(
	resolved: ResolvedCustomization,
	id: RibbonControlId,
): boolean {
	if (resolved.hiddenRibbonControls.has(id)) {
		return false;
	}
	for (const [alias, covered] of Object.entries(RIBBON_CONTROL_ALIASES)) {
		if (
			covered.includes(id) &&
			((isRibbonControlId(alias) && resolved.hiddenRibbonControls.has(alias)) ||
				(isRibbonGroupId(alias) && resolved.hiddenRibbonGroups.has(alias)))
		) {
			return false;
		}
	}
	const group = id.slice(0, id.lastIndexOf('.')) as RibbonGroupId;
	return !resolved.hiddenRibbonGroups.has(group);
}

/** The attribute map a binding spreads onto a group wrapper. */
export function ribbonGroupAttrs(id: RibbonGroupId): Record<typeof RIBBON_GROUP_ATTR, string> {
	return { [RIBBON_GROUP_ATTR]: id };
}

/** The attribute map a binding spreads onto a control. */
export function ribbonControlAttrs(
	id: RibbonControlId,
): Record<typeof RIBBON_CONTROL_ATTR, string> {
	return { [RIBBON_CONTROL_ATTR]: id };
}

function attrSelector(scope: string | undefined, attr: string, id: string): string {
	const prefix = scope ? `[${RIBBON_SCOPE_ATTR}="${scope}"] ` : '';
	return `${prefix}[${attr}="${id}"]`;
}

/**
 * The stylesheet that removes every group and control the host hid, or `''`
 * when nothing is hidden. Pass the viewer's scope token (the value of its
 * {@link RIBBON_SCOPE_ATTR}) to confine the rules to that viewer; omit it for
 * a document-wide sheet. Ids are validated against the catalogue, so a stray
 * string from a host cannot inject CSS.
 */
export function ribbonCustomizationCss(resolved: ResolvedCustomization, scope?: string): string {
	const safeScope = scope && /^[\w-]+$/u.test(scope) ? scope : undefined;
	const selectors: string[] = [];
	for (const id of resolved.hiddenInspectorCards) {
		if (INSPECTOR_CARD_IDS.includes(id)) {
			selectors.push(attrSelector(safeScope, 'data-pptx-inspector-card', id));
		}
	}
	for (const id of resolved.hiddenFormatSections) {
		if (FORMAT_SECTION_IDS.includes(id)) {
			selectors.push(attrSelector(safeScope, FORMAT_SECTION_ATTRIBUTE, id));
		}
	}
	for (const id of resolved.hiddenRibbonGroups) {
		if (isRibbonGroupId(id)) {
			selectors.push(attrSelector(safeScope, RIBBON_GROUP_ATTR, id));
			for (const covered of RIBBON_CONTROL_ALIASES[id] ?? []) {
				selectors.push(attrSelector(safeScope, RIBBON_CONTROL_ATTR, covered));
			}
		}
	}
	for (const id of resolved.hiddenRibbonControls) {
		if (isRibbonControlId(id)) {
			selectors.push(attrSelector(safeScope, RIBBON_CONTROL_ATTR, id));
			if (id.startsWith(`${SMART_ART_ARRANGE_GROUP}.`)) {
				// The embedded Arrange strip keeps its Home IDs in every binding.
				const group = attrSelector(safeScope, RIBBON_GROUP_ATTR, SMART_ART_ARRANGE_GROUP);
				const homeId = `home.arrange${id.slice(SMART_ART_ARRANGE_GROUP.length)}`;
				selectors.push(`${group} ${attrSelector(undefined, RIBBON_CONTROL_ATTR, homeId)}`);
			}
			for (const covered of RIBBON_CONTROL_ALIASES[id] ?? []) {
				selectors.push(attrSelector(safeScope, RIBBON_CONTROL_ATTR, covered));
			}
		}
	}
	const order = formatSectionOrderCss(resolved, safeScope);
	if (selectors.length === 0) {
		return order;
	}
	return `${selectors.join(',\n')} {\n\tdisplay: none !important;\n}\n${order}`;
}

/**
 * Move the host's `inspector.sectionOrder` sections to the front of the
 * Format pane: the section list is a flex column, so a negative `order`
 * places a listed section ahead of the rest, which keep their own order.
 */
function formatSectionOrderCss(resolved: ResolvedCustomization, scope: string | undefined): string {
	const ids = resolved.formatSectionOrder.filter((id) => FORMAT_SECTION_IDS.includes(id));
	return ids
		.map(
			(id, index) =>
				`${attrSelector(scope, FORMAT_SECTION_ATTRIBUTE, id)} {\n\torder: ${index - ids.length};\n}\n`,
		)
		.join('');
}
