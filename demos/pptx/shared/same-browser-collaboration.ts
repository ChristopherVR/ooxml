/**
 * `?sameBrowser=<room>`: the shared-contract `collaboration` input with only a room and a user,
 * which every binding opens as a BroadcastChannel room between windows of one browser. Add
 * `&name=<user>` to name the user and `&join=1` to join (never seed) the room.
 */
interface SameBrowserState {
	status: string;
	synced: boolean;
	people: readonly { name: string; self: boolean }[];
}

export interface SameBrowserDemo {
	/** Assignable to every binding's `collaboration` prop. */
	collaboration: { roomId: string; user: { name: string }; intent: 'create' | 'join' };
	source: Uint8Array;
	fileName: string;
	/** Feed the binding's collaboration-change event here; it fills the status readout. */
	report(state: SameBrowserState): void;
	dispose(): void;
}

export const sameBrowserRoom = (): string | null =>
	new URLSearchParams(location.search).get('sameBrowser') || null;

export async function createSameBrowserDemo(): Promise<SameBrowserDemo> {
	const params = new URLSearchParams(location.search);
	const response = await fetch(`${import.meta.env.BASE_URL}sample-deck.pptx`);
	if (!response.ok) {
		throw new Error(`Sample load failed: ${response.status}`);
	}
	const source = new Uint8Array(await response.arrayBuffer());
	const output = document.createElement('output');
	output.setAttribute('aria-label', 'Same-browser collaboration');
	output.style.cssText =
		'position:fixed;inset:0 0 auto;z-index:100000;padding:4px 12px;background:#f8fafc;color:#0f172a;font:13px system-ui;height:28px;box-sizing:border-box';
	output.textContent = 'Same browser: starting';
	document.body.append(output);
	return {
		collaboration: {
			roomId: params.get('sameBrowser') ?? '',
			user: { name: params.get('name') || 'Same-browser user' },
			intent: params.get('join') === '1' ? 'join' : 'create',
		},
		source,
		fileName: 'sample-deck.pptx',
		report(state) {
			const names = state.people.map((p) => (p.self ? `${p.name} (you)` : p.name)).join(', ');
			output.textContent = `Same browser: ${state.status}; synced: ${state.synced}; people: ${state.people.length}; ${names}`;
		},
		dispose() {
			output.remove();
		},
	};
}
