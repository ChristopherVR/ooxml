import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import type { AddInDialogHost } from './add-in-dialog';

/** The host's profile-scoped preference store. Bytes are always revalidated when loaded. */
export interface OfficeAddInPreferences {
	read(host: OfficeAddInHost): Promise<readonly Uint8Array[]>;
	write(host: OfficeAddInHost, manifestId: string, bytes: Uint8Array | undefined): Promise<void>;
}

export const ADD_IN_PREFERENCES_CHANGED = 'office-add-in-preferences-changed';
const PREFERENCES = Symbol.for('office-ui.add-in-preferences');
type Host = AddInDialogHost & { [PREFERENCES]?: OfficeAddInPreferences | null };

export function setOfficeAddInPreferences(host: Host, preferences: OfficeAddInPreferences | null) {
	if (host[PREFERENCES] === preferences) return;
	host[PREFERENCES] = preferences;
	host.dispatchEvent(new Event(ADD_IN_PREFERENCES_CHANGED));
}

export function addInPreferences(scopes: readonly Host[]) {
	for (const scope of scopes) if (scope[PREFERENCES] !== undefined) return scope[PREFERENCES];
}
