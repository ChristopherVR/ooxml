import type { LocalAddIn } from 'ooxml-core/automation/add-ins';
import type { OpenAddInsOptions } from '../ribbon/add-in-dialog';
import { editorAddIns } from '../ribbon/add-in-installed-tabs';
import { addInLauncher } from '../ribbon/add-in-launch';
import { INSTALLED_ADD_IN_LIMIT, localAddIns } from '../ribbon/add-in-installations';
import { addInLocale, localAddInText } from '../ribbon/add-in-local-strings';
import { addInManifestTabs } from '../ribbon/add-in-manifest-tabs';
import { addInCapabilities, isSelectionSet } from '../ribbon/add-in-requirements';
import { addInRuntime } from '../ribbon/add-in-runtime-settings';
import { addInPreferences } from '../ribbon/add-in-preferences';
import { addInSettings } from '../ribbon/add-in-settings';

const sourceBytes = new WeakMap<LocalAddIn, Uint8Array>();
const validatedWithSelection = new WeakMap<LocalAddIn, boolean>();

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
	const hasSelection = () =>
		product === 'word' ||
		product === 'excel' ||
		Boolean((addInRuntime(scopes)?.selection ?? options.selection)?.());
	const manifestOptions = (selection = hasSelection()) => ({
		locale: manifestLocale,
		capabilities: addInCapabilities(selection),
	});
	const syncTabs = () => addInManifestTabs(host, installed, launcher, t);
	const restore =
		store && product !== 'visio'
			? async (signal: AbortSignal) => {
					const files = await store.read(product);
					if (signal.aborted) return;
					const restored = new Map<string, LocalAddIn>();
					if (files.length) {
						const { readLocalAddIn } = await import('../ribbon/add-in-manifest-file');
						const selection = hasSelection();
						for (const file of files.slice(0, INSTALLED_ADD_IN_LIMIT)) {
							if (signal.aborted) return;
							try {
								const bytes = new Uint8Array(file);
								const addIn = await readLocalAddIn(bytes, {
									...manifestOptions(selection),
									host: product,
								});
								const prior = installed.get(addIn.manifest.id);
								const entry =
									prior?.digest === addIn.digest &&
									prior.manifest.name === addIn.manifest.name &&
									validatedWithSelection.get(prior) === selection
										? prior
										: addIn;
								sourceBytes.set(entry, bytes);
								validatedWithSelection.set(entry, selection);
								restored.set(addIn.manifest.id, entry);
							} catch {
								// A stale or corrupt record cannot grant access or hide other installs.
							}
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
		host,
		installed,
		launcher,
		syncTabs,
		restore,
		tabs: editorAddIns(host),
		t,
		product,
		locale,
		manifestOptions,
		async refreshRuntime(signal: AbortSignal) {
			if (product === 'powerpoint') {
				const selection = hasSelection();
				for (const [id, entry] of installed) {
					const bytes = sourceBytes.get(entry);
					if (!bytes || validatedWithSelection.get(entry) === selection) continue;
					if (
						!entry.manifest.requirements.some(
							({ kind, name }) => kind === 'set' && isSelectionSet(name),
						)
					)
						continue;
					const { readLocalAddIn } = await import('../ribbon/add-in-manifest-file');
					const refreshed = await readLocalAddIn(bytes, {
						...manifestOptions(selection),
						host: product,
					});
					if (signal.aborted) return;
					if (installed.get(id) !== entry || refreshed.digest !== entry.digest) continue;
					sourceBytes.set(refreshed, bytes);
					validatedWithSelection.set(refreshed, selection);
					installed.set(id, refreshed);
				}
			}
			if (!signal.aborted) syncTabs();
		},
		restoreFocus,
		async install({ addIn, bytes, previous }: AddInPreview) {
			const id = addIn.manifest.id;
			if (installed.get(id) !== previous || (!previous && installed.size >= INSTALLED_ADD_IN_LIMIT))
				throw new Error('Add-in installation changed or reached its limit.');
			const owned = new Uint8Array(bytes);
			if (store && product !== 'visio') await store.write(product, id, bytes);
			if (addInPreferences(scopes) !== store) return;
			launcher.close(id);
			sourceBytes.set(addIn, owned);
			validatedWithSelection.set(addIn, hasSelection());
			installed.set(id, addIn);
			await syncTabs();
		},
		async remove(id: string) {
			if (store && product !== 'visio') await store.write(product, id, undefined);
			if (addInPreferences(scopes) !== store) return;
			launcher.close(id);
			installed.delete(id);
			syncTabs();
		},
	};
}

export type AddInSession = ReturnType<typeof addInSession>;
