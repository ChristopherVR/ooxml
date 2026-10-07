/** Stable stop ordering shared by gradient authoring in every Office format. */
export function sortGradientStops<T extends { position: number }>(stops: readonly T[]): T[] {
	return stops.slice().sort((a, b) => a.position - b.position);
}

export function updateGradientStop<T extends { position: number }>(
	stops: readonly T[],
	index: number,
	changes: Partial<T>,
): T[] {
	return sortGradientStops(stops.map((stop, i) => (i === index ? { ...stop, ...changes } : stop)));
}

/** Keep two stops, matching the Office gradient editor's minimum. */
export function removeGradientStop<T extends { position: number }>(
	stops: readonly T[],
	index: number,
): T[] | undefined {
	return stops.length <= 2 ? undefined : stops.filter((_, i) => i !== index);
}
