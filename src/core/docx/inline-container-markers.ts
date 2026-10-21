// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Markers read inside a content control or simple field that a rebuilt paragraph continues by its
// identity go with it, mapped through a character diff of the container's own text.
import type { ContainerSource } from './inline-containers';
import { moveMarkers, positionText, type MarkerAt } from './inline-marker-positions';
import type { TextRun } from './model';
import type { XmlElement } from './xml';

// Stands beside a container's text where the container did not touch the paragraph's edge, so a
// marker at the container's edge is not treated as one at the paragraph's.
const GUARD: TextRun = { text: '\uFFFF' };

/**
 * Moves the markers read inside a source container matched by its identity (`follow`) into the
 * group that continues it, mapped through a character diff of the container's own text, so they
 * go with it.
 */
export function followContainers<T extends MarkerAt & { source: number; container?: XmlElement }>(
	markers: readonly T[],
	sources: ReadonlyMap<number, ContainerSource>,
	before: readonly TextRun[],
	after: readonly TextRun[],
): T[] {
	const continued = new Map<XmlElement, ContainerSource | undefined>();
	for (const source of sources.values())
		for (const container of source.containers)
			if (container && source.same)
				continued.set(container, continued.has(container) ? undefined : source);
	const length = before.map(positionText).join('').length;
	// A container's markers are mapped together, so they share each placement the diff makes.
	const groups = new Map<ContainerSource, T[]>();
	for (const marker of markers) {
		const source = marker.container && continued.get(marker.container);
		if (source?.follow) groups.set(source, [...(groups.get(source) ?? []), marker]);
	}
	const placed = new Map<T, number>();
	for (const [source, group] of groups) {
		const lead = source.from.start > 0 ? [GUARD] : [];
		const trail = source.from.end < length ? [GUARD] : [];
		const moved = moveMarkers(
			[...lead, ...before.slice(source.fromRuns.start, source.fromRuns.end), ...trail],
			[...lead, ...after.slice(source.toRuns.start, source.toRuns.end), ...trail],
			group.map((marker) => ({
				...marker,
				position: marker.source - source.from.start + lead.length,
			})),
		);
		group.forEach((marker, index) =>
			placed.set(marker, source.to.start + moved[index]!.position - lead.length),
		);
	}
	return markers.map((marker) =>
		placed.has(marker) ? { ...marker, position: placed.get(marker)! } : marker,
	);
}
