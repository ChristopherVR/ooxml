const selectionSets = new Map([
	['selection', '1.1'],
	['textcoercion', '1.1'],
]);

export function isSelectionSet(name: string): boolean {
	return selectionSets.has(name.toLowerCase());
}

export function addInCapabilities(hasSelection: boolean) {
	return hasSelection ? { sets: selectionSets } : {};
}

export function isAddInSetSupported(hasSelection: boolean, name: string, minVersion = '1.1') {
	if (!hasSelection || typeof name !== 'string' || typeof minVersion !== 'string') return false;
	if (!isSelectionSet(name) || !/^\d{1,5}(?:\.\d{1,5}){0,3}$/.test(minVersion)) return false;
	const [major = 0, minor = 0, ...rest] = minVersion.split('.').map(Number);
	return (
		major < 1 || (major === 1 && (minor < 1 || (minor === 1 && rest.every((part) => part === 0))))
	);
}
