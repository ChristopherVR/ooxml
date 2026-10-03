/** Published artifacts may use registry dependencies, never private source packages. */
export function forbiddenManifestEntries(manifest) {
	const found = [];
	for (const field of [
		'dependencies',
		'peerDependencies',
		'optionalDependencies',
		'devDependencies',
	]) {
		for (const name of Object.keys(manifest[field] ?? {})) {
			if (
				[
					'@christophervr/visio-viewer',
					'@christophervr/visio-viewer-bindings',
					'@christophervr/ole2',
				].includes(name)
			)
				found.push(field + '.' + name);
		}
	}
	return found;
}
