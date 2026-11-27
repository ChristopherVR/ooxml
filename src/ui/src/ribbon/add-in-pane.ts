import type { OfficeAddInHost } from 'ooxml-core/automation/add-ins';
import { html, nothing } from 'lit';
import { controlStyles, OfficeElement } from '../base';
import { defineTaskPane, type OfficeUiTaskPane } from '../chrome/task-pane';
import { definer } from '../registry';
import type { OfficeServiceText } from '../dialog/office-service-dialog';
import type { AddInDialogHost } from './add-in-dialog';
import { ADD_IN_RUNTIME_CHANGED, type OfficeAddInSelection } from './add-in-runtime-settings';
import {
	ADD_IN_TIMEOUT,
	HOST_NAMES,
	bridgeRequest,
	dispatchAddInRequest,
	isPaneHandshake,
	type BridgeResult,
} from './add-in-protocol';
import { AddInPaneWidth } from './add-in-pane-resize';
import css from './add-in-pane.css?raw';

interface PaneOptions {
	id: string;
	url: URL;
	name: string;
	product: OfficeAddInHost;
	permissions: readonly string[];
	selection: OfficeAddInSelection | undefined;
	active: () => boolean;
	scopes: readonly AddInDialogHost[];
	t: OfficeServiceText;
	restoreFocus: () => void;
}

class OfficeUiAddInPane extends OfficeElement {
	static override styles = controlStyles(css);

	#options: PaneOptions | undefined;
	#frame: HTMLIFrameElement | undefined;
	#status = '';
	#width = new AddInPaneWidth(this, () => this.requestUpdate());
	#sizeObserver: ResizeObserver | undefined;
	#stop: (() => void) | undefined;

	show(options: PaneOptions): void {
		this.#options = options;
		const window = this.ownerDocument.defaultView;
		if (!window) return;

		const dock = this.parentElement;
		if (dock && window.ResizeObserver) {
			this.#sizeObserver = new window.ResizeObserver(() => this.#width.layout());
			this.#sizeObserver.observe(dock);
		}
		this.#width.layout();

		const frame = this.ownerDocument.createElement('iframe');
		frame.title = options.name;
		frame.setAttribute('sandbox', 'allow-scripts');
		frame.setAttribute(
			'allow',
			"camera 'none'; microphone 'none'; geolocation 'none'; clipboard-read 'none'; clipboard-write 'none'; payment 'none'; usb 'none'",
		);
		frame.referrerPolicy = 'no-referrer';
		frame.setAttribute('credentialless', '');
		frame.hidden = true;
		let port: MessagePort | undefined;
		let loaded = false;
		let calls = 0;
		let lastId = 0;
		const detach = () => window.removeEventListener('message', connect);
		const timer = window.setTimeout(() => {
			detach();
			frame.remove();
			this.#status = options.t('The add-in did not connect to this editor.');
			this.requestUpdate();
		}, ADD_IN_TIMEOUT);
		const connect = (event: MessageEvent) => {
			if (!options.active() || !isPaneHandshake(event, frame.contentWindow)) return;
			port = event.ports[0];
			if (!port) return;
			detach();
			window.clearTimeout(timer);
			const channel = port;
			channel.onmessage = (message: MessageEvent<unknown>) => {
				const request = bridgeRequest(message.data);
				if (!request || request.id <= lastId || !this.isConnected) return;
				lastId = request.id;
				let result: BridgeResult;
				try {
					result = !options.active()
						? { error: 'This add-in no longer has permission to run.' }
						: ++calls <= 256
							? dispatchAddInRequest(request, options.permissions, options.selection)
							: { error: 'This task pane reached its request limit. Reopen it to continue.' };
				} catch {
					result = { error: 'The editor could not handle this selection request.' };
				}
				channel.postMessage({ id: request.id, ...result });
				if (calls > 256) channel.close();
			};
			channel.postMessage({
				kind: 'ready',
				host: HOST_NAMES[options.product],
				hasSelection: Boolean(options.selection),
			});
			frame.hidden = false;
			this.#status = '';
			this.requestUpdate();
		};
		const revoke = () => {
			if (options.active()) return;
			this.remove();
			options.restoreFocus();
		};
		for (const scope of options.scopes) scope.addEventListener(ADD_IN_RUNTIME_CHANGED, revoke);
		this.#stop = () => {
			detach();
			for (const scope of options.scopes) scope.removeEventListener(ADD_IN_RUNTIME_CHANGED, revoke);
			window.clearTimeout(timer);
			port?.close();
			frame.remove();
		};
		frame.addEventListener('load', () => {
			if (loaded) {
				this.#stop?.();
				this.#stop = undefined;
				this.#status = options.t('The add-in navigated away and was disconnected.');
				this.requestUpdate();
			}
			loaded = true;
		});
		window.addEventListener('message', connect);
		this.#status = options.t('Loading add-ins…');
		frame.src = options.url.href;
		this.#frame = frame;
		this.requestUpdate();
	}

	override connectedCallback() {
		super.connectedCallback();
		if (this.#options && !this.#stop) this.remove();
	}

	override disconnectedCallback() {
		this.#sizeObserver?.disconnect();
		this.#stop?.();
		this.#stop = undefined;
		super.disconnectedCallback();
	}

	protected override render() {
		const options = this.#options;
		if (!options) return nothing;
		const close = () => {
			this.remove();
			options.restoreFocus();
		};
		const width = this.#width;
		return html`<div
				class="resize"
				role="separator"
				tabindex="0"
				aria-label=${options.t('Resize add-in pane')}
				aria-controls="add-in-content"
				aria-orientation="vertical"
				aria-valuemin=${width.min}
				aria-valuemax=${width.max}
				aria-valuenow=${width.value}
				@keydown=${width.key}
				@pointerdown=${width.press}
				@pointermove=${width.move}
				@lostpointercapture=${width.release}
			></div
			><office-ui-task-pane
				id="add-in-content"
				.label=${options.name}
				.closeLabel=${options.t('Close')}
				@office-pane-close=${close}
			>
				<p role="status" ?hidden=${!this.#status}>${this.#status}</p>
				${this.#frame}
			</office-ui-task-pane>`;
	}
}

const define = definer('office-ui-add-in-pane', () => OfficeUiAddInPane, [defineTaskPane]);

export function openAddInPane(host: AddInDialogHost, options: PaneOptions): void {
	define(host.ownerDocument.defaultView?.customElements);
	host.querySelector('office-ui-add-in-pane')?.remove();
	const dock = host.querySelector('[data-office-add-in-dock], [data-pptx-chrome="body"]') ?? host;
	const pane = host.ownerDocument.createElement('office-ui-add-in-pane');
	if (!('show' in pane) || typeof pane.show !== 'function') return;
	pane.dataset.addInId = options.id;
	dock.append(pane);
	pane.show(options);
	pane.shadowRoot?.querySelector<OfficeUiTaskPane>('office-ui-task-pane')?.focusHeading();
}
