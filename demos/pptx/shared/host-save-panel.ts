/**
 * A "host save" panel for the demos: what a host application does with the revision-aware
 * save contract (`getSaveSnapshot`, `acknowledgeSave`), with a controllable persistence delay
 * and failure so the behaviour can be seen and tested in every binding.
 *
 * It is framework-free DOM, shown only when the page URL carries `?hostSave=1`, and reads the
 * viewer through the same handle the demos already expose for development.
 */

interface SaveHandle {
	getSaveRevision(): number;
	getSaveSnapshot(): Promise<{ revision: number; content: Uint8Array }>;
	acknowledgeSave(revision: number): { status: string };
	isDirty(): boolean;
	canUndo(): boolean;
	undo(): void;
	getSlides(): { id: string; elements: { id: string; x: number }[] }[];
	updateElement(elementId: string, patch: { x: number }): void;
}

interface HostSaveWindow extends Window {
	/** The snapshots the panel has taken, oldest first (test and debugging aid). */
	__hostSaveSnapshots?: { revision: number; content: Uint8Array }[];
}

const STYLE =
	'position:fixed;right:8px;bottom:8px;z-index:2147483000;display:flex;flex-direction:column;' +
	'gap:4px;padding:8px;font:12px/1.3 system-ui,sans-serif;background:#fffd;color:#111;' +
	'border:1px solid #888;border-radius:6px;max-width:260px';

function element<K extends keyof HTMLElementTagNameMap>(
	tag: K,
	props: Partial<HTMLElementTagNameMap[K]> & { testId?: string },
): HTMLElementTagNameMap[K] {
	const { testId, ...rest } = props;
	const node = Object.assign(document.createElement(tag), rest);
	if (testId) {
		node.dataset.testid = testId;
	}
	return node;
}

/** Mount the panel once; it asks `getHandle` for the viewer whenever a button is pressed. */
export function mountHostSavePanel(getHandle: () => unknown): () => void {
	const host = window as HostSaveWindow;
	host.__hostSaveSnapshots = [];
	const viewer = () => getHandle() as SaveHandle | null | undefined;

	const root = element('aside', { testId: 'host-save-panel', ariaLabel: 'Host save' });
	root.style.cssText = STYLE;
	const delay = element('input', { type: 'number', value: '300', testId: 'host-save-delay' });
	delay.style.width = '70px';
	const fail = element('input', { type: 'checkbox', testId: 'host-save-fail' });
	const save = element('button', { textContent: 'Host save', testId: 'host-save' });
	const edit = element('button', { textContent: 'Host edit', testId: 'host-save-edit' });
	const undo = element('button', { textContent: 'Host undo', testId: 'host-save-undo' });
	const ack = element('button', { textContent: 'Acknowledge', testId: 'host-save-ack' });
	const ackRevision = element('input', {
		type: 'number',
		value: '1',
		testId: 'host-save-ack-revision',
	});
	ackRevision.style.width = '50px';
	const inFlight = element('output', { textContent: '0', testId: 'host-save-inflight' });
	const result = element('output', { textContent: '', testId: 'host-save-result' });
	const state = element('output', { textContent: '', testId: 'host-save-state' });
	const log = element('ol', { testId: 'host-save-log' });
	log.style.cssText = 'margin:0;padding-left:16px;max-height:90px;overflow:auto';

	const row = (label: string, ...nodes: Node[]) => {
		const line = document.createElement('label');
		line.style.cssText = 'display:flex;gap:4px;align-items:center';
		line.append(label, ...nodes);
		return line;
	};
	root.append(
		row('Delay ms', delay),
		row('Fail', fail),
		save,
		edit,
		undo,
		row('Ack rev', ackRevision, ack),
		row('In flight', inFlight),
		row('Last', result),
		row('State', state),
		log,
	);

	const refresh = () => {
		const api = viewer();
		state.textContent = api
			? JSON.stringify({
					dirty: api.isDirty(),
					revision: api.getSaveRevision(),
					canUndo: api.canUndo(),
				})
			: '';
	};
	const report = (text: string) => {
		result.textContent = text;
		log.append(element('li', { textContent: text }));
		refresh();
	};

	let pending = 0;
	save.addEventListener('click', async () => {
		const api = viewer();
		if (!api) {
			return;
		}
		const failing = fail.checked;
		const wait = Number(delay.value) || 0;
		pending += 1;
		inFlight.textContent = String(pending);
		try {
			const snapshot = await api.getSaveSnapshot();
			host.__hostSaveSnapshots?.push(snapshot);
			await new Promise((done) => setTimeout(done, wait));
			if (failing) {
				throw new Error('persistence failed');
			}
			const outcome = api.acknowledgeSave(snapshot.revision);
			report(`${outcome.status} ${snapshot.revision}`);
		} catch (error) {
			report(`failed ${String(error)}`);
		} finally {
			pending -= 1;
			inFlight.textContent = String(pending);
		}
	});
	ack.addEventListener('click', () => {
		const api = viewer();
		if (api) {
			const revision = Number(ackRevision.value);
			report(`${api.acknowledgeSave(revision).status} ${revision}`);
		}
	});
	edit.addEventListener('click', () => {
		const api = viewer();
		const first = api?.getSlides()[0]?.elements[0];
		if (api && first) {
			api.updateElement(first.id, { x: first.x + 10 });
			refresh();
		}
	});
	undo.addEventListener('click', () => {
		viewer()?.undo();
		refresh();
	});

	document.body.append(root);
	const timer = window.setInterval(refresh, 250);
	return () => {
		window.clearInterval(timer);
		root.remove();
		delete host.__hostSaveSnapshots;
	};
}

/** Whether the page asked for the panel. */
export function hostSaveRequested(): boolean {
	return new URLSearchParams(window.location.search).get('hostSave') === '1';
}
