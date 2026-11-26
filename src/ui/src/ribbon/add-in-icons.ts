import type { AddInManifest } from 'ooxml-core/automation/add-ins';

export function addInIcon(manifest: AddInManifest): string | undefined {
	for (const icon of manifest.icons)
		if (icon.kind === 'package' && icon.preview) return icon.preview;
}
