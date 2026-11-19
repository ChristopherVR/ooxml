/**
 * The demos keep the AI settings a developer sets by hand for the current tab only, in
 * sessionStorage, as the suite keeps AI keys, and delete what an earlier demo version told them to
 * leave in localStorage: no key outlives the tab on the shared site origin.
 */
export function readDemoAiSetting(key: string): string | null {
	try {
		localStorage.removeItem(key);
		return sessionStorage.getItem(key);
	} catch {
		return null;
	}
}
