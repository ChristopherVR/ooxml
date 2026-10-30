import { attribute } from './sheet';
import type { ConnectorSite, End } from './edit-connector-layout';
import type { ConnectorGlue } from './edit-connector';
import { topShape } from './edit-connector';
import { effectiveCells, effectiveNumber } from './edit-stencil-connector';

/** ShapeRouteStyle of the Organization Chart stencil's Dynamic connector (visLORouteOrgChartNS). */
const ORG_CHART_ROUTE = 3;

/**
 * The sides an org chart reporting line joins, or nothing for any other connector. A Dynamic
 * connector whose route style is the organization chart's runs from the bottom of the manager to
 * the top of the subordinate, as Visio's org chart routing draws it; to an Assistant Belt it
 * drops from the manager's bottom and enters the assistant's side nearest that line. Only ends
 * glued to whole shapes are handled; an end glued to a connection point keeps its point.
 */
export function orgChartSites(
	root: Element,
	connector: Element,
	glue: ConnectorGlue,
	sites: (end: End) => ConnectorSite[],
): Record<End, ConnectorSite> | undefined {
	const begin = glue.ends.begin;
	const end = glue.ends.end;
	if (!begin || !end || begin.point !== undefined || end.point !== undefined) return undefined;
	if (effectiveNumber(effectiveCells(connector), 'ShapeRouteStyle', 0) !== ORG_CHART_ROUTE)
		return undefined;
	const from = sites('begin');
	const to = sites('end');
	if (from.length !== 4 || to.length !== 4) return undefined;
	const lowest = (list: ConnectorSite[]) =>
		list.reduce((best, site) => (site.point.y < best.point.y ? site : best));
	const highest = (list: ConnectorSite[]) =>
		list.reduce((best, site) => (site.point.y > best.point.y ? site : best));
	const manager = lowest(from);
	const assistant = /^Assistant Belt/i.test(attribute(topShape(root, end.target), 'NameU') ?? '');
	if (!assistant) return { begin: manager, end: highest(to) };
	// The side of the assistant that faces the manager's line.
	const top = highest(to).point.y;
	const bottom = lowest(to).point.y;
	const sideways = to.filter((site) => site.point.y < top - 1e-9 && site.point.y > bottom + 1e-9);
	if (!sideways.length) return undefined;
	const facing = sideways.reduce((best, site) =>
		Math.abs(site.point.x - manager.point.x) < Math.abs(best.point.x - manager.point.x)
			? site
			: best,
	);
	return { begin: manager, end: facing };
}
