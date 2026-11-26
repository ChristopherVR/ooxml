import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OpenAddInsOptions } from '../ribbon/add-in-dialog';
import { editorAddIns } from '../ribbon/add-in-installed-tabs';
import { addInLauncher } from '../ribbon/add-in-launch';
import { INSTALLED_ADD_IN_LIMIT, localAddIns } from '../ribbon/add-in-local-dialog';
import { addInLocale, localAddInText } from '../ribbon/add-in-local-strings';
import { addInManifestTabs } from '../ribbon/add-in-manifest-tabs';
import { addInCapabilities } from '../ribbon/add-in-requirements';
import { addInRuntime } from '../ribbon/add-in-runtime-settings';
import { addInPreferences } from '../ribbon/add-in-preferences';
import { addInSettings } from '../ribbon/add-in-settings';

export interface AddInPreview {
	addIn: LocalAddIn;
	bytes: Uint8Array;
	previous: LocalAddIn | undefined;
}

export function addInSession(options: OpenAddInsOptions, restoreFocus: () => void) {
	const { host, manifestLocale, product, scopes } = options;
	const locale = addInLocale(manifestLocale);
	const installed = localAddIns(host);
	const t = localAddInText(locale, options.t);
	const launcher = addInLauncher(options, installed, t);
	const store = addInPreferences(scopes);
	const manifestOptions = () => ({
		locale: manifestLocale,
		capabilities: addInCapabilities(
			product === 'word' ||
				product === 'excel' ||
				Boolean(addInRuntime(scopes)?.selection ?? options.selection),
		),
	});
	const syncTabs = () => addInManifestTabs(host, installed, launcher, t);
	const restore =
		store && product !== 'visio'
			? async (signal: AbortSignal) => {
					const files = await store.read(product);
					if (signal.aborted || (!files.length && !installed.size)) return;
					const { readLocalAddIn } = await import('../ribbon/add-in-manifest-file');
					const restored = new Map<string, LocalAddIn>();
					for (const file of files.slice(0, INSTALLED_ADD_IN_LIMIT)) {
						if (signal.aborted) return;
						try {
							const addIn = await readLocalAddIn(file, { ...manifestOptions(), host: product });
							const prior = installed.get(addIn.manifest.id);
							restored.set(
								addIn.manifest.id,
								prior?.digest === addIn.digest && prior.manifest.name === addIn.manifest.name
									? prior
									: addIn,
							);
						} catch {
							// A stale or corrupt record cannot grant access or hide other installs.
						}
					}
					if (signal.aborted) return;
					for (const [id, addIn] of installed) if (restored.get(id) !== addIn) launcher.close(id);
					installed.clear();
					for (const [id, addIn] of restored) installed.set(id, addIn);
				}
			: undefined;

	return {
		...addInSettings(scopes),
		installed,
		launcher,
		syncTabs,
		restore,
		tabs: editorAddIns(host),
		t,
		product,
		locale,
		manifestOptions,
		restoreFocus,
		async install({ addIn, bytes, previous }: AddInPreview) {
			const id = addIn.manifest.id;
			if (installed.get(id) !== previous || (!previous && installed.size >= INSTALLED_ADD_IN_LIMIT))
				throw new Error('Add-in installation changed or reached its limit.');
			if (store && product !== 'visio') await store.write(product, id, bytes);
			launcher.close(id);
			installed.set(id, addIn);
			syncTabs();
		},
		async remove(id: string) {
			if (store && product !== 'visio') await store.write(product, id, undefined);
			launcher.close(id);
			installed.delete(id);
			syncTabs();
		},
	};
}

export type AddInSession = ReturnType<typeof addInSession>;
