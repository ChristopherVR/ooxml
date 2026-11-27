/**
 * The demos' shared add-in preference store name and off switch (`?addInsStorage=off`). Each
 * demo passes `createAddInPreferencesStore` from `ooxml-ui/suite`, which its own package declares.
 */
export function demoAddInPreferences<Store>(create: (name: string) => Store): Store | null {
	if (new URLSearchParams(location.search).get('addInsStorage') === 'off') return null;
	return create('ooxml-demo-add-ins');
}
