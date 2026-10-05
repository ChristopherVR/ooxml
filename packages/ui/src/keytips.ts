import { tok } from './tokens.js';
/**
 * Office KeyTips: press and release Alt to show letter badges on ribbon controls, then type a
 * badge to activate it. Products annotate controls only:
 * - `data-keytip="H"` marks a control (one or two characters);
 * - `data-keytip-panel="home-panel"` on a control (a ribbon tab) makes the controls inside the
 *   element with that id the next level after it is activated.
 * First-level tips are those not inside any `[data-keytip-level]` container; a panel element
 * carries `data-keytip-level` so its tips only appear after its tab. Escape steps back a level;
 * Alt, a click or focus leaving the scope exits. Disabled controls show dimmed and never run.
 * Badges render in a `popover` top layer, so containment and overflow never misplace them.
 */
export interface KeyTipsHandle {
	/** Whether badges are currently shown. */
	readonly active: boolean;
	/** Show the first level, as Alt does; for products that also start KeyTips from F10. */
	start(): void;
	/** Hide the badges. */
	stop(): void;
	dispose(): void;
}

const BADGE_CSS =
	'position:fixed;inset:auto;margin:0;padding:0;border:0;background:transparent;overflow:visible;pointer-events:none';
const TIP_CSS =
	`position:fixed;transform:translate(-50%,-30%);min-width:${tok('--office-icon-size')};` +
	`padding:${tok('--office-space-px')} ${tok('--office-space-1')};border:${tok('--office-border-width')} solid ${tok('--office-keytip-border')};` +
	`border-radius:${tok('--office-radius-xs')};background:${tok('--office-keytip-background')};color:${tok('--office-keytip-foreground')};` +
	`font:${tok('--office-font-weight-bold')} ${tok('--office-font-size-xs')}/${tok('--office-icon-size')} ${tok('--office-font')};` +
	`text-align:center;box-shadow:${tok('--office-shadow-sm')}`;

type Root = ShadowRoot | HTMLElement;
const isDisabled = (el: Element) =>
	el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true';
const visible = (el: Element) => {
	const box = el.getBoundingClientRect();
	return box.width > 0 && box.height > 0;
};

/**
 * Marked controls under `root`, including inside open shadow roots of shared elements (the
 * ribbon renders its tabs in its own shadow root). A shadow-hosted control has no light-DOM
 * `[data-keytip-level]` ancestor, so it belongs to the level of its host.
 */
function deepTips(root: ParentNode): HTMLElement[] {
	const found = [...root.querySelectorAll<HTMLElement>('[data-keytip]')];
	for (const host of root.querySelectorAll('*')) {
		const shadow = (host as Element).shadowRoot;
		if (!shadow) continue;
		const inner = deepTips(shadow);
		if (!inner.length) continue;
		// Treat inner controls as if they sat where their host is, for level ownership.
		for (const el of inner) hostOf.set(el, host);
		found.push(...inner);
	}
	return found;
}
const hostOf = new WeakMap<Element, Element>();

/** Activate a control as a click would, reaching into shared controls' shadow buttons. */
function activate(el: HTMLElement): void {
	const inner = el.shadowRoot?.querySelector<HTMLElement>('.main, button');
	(inner ?? el).click();
}

export function attachKeyTips(scope: Root): KeyTipsHandle {
	const doc = (scope instanceof ShadowRoot ? scope.host : scope).ownerDocument;
	const view = doc.defaultView;
	const layer = doc.createElement('div');
	layer.setAttribute('aria-hidden', 'true');
	layer.style.cssText = BADGE_CSS;
	const topLayer = typeof (layer as { showPopover?: unknown }).showPopover === 'function';
	if (topLayer) layer.setAttribute('popover', 'manual');
	else layer.hidden = true;
	scope.append(layer);
	let levels: Element[] = [];
	let typed = '';
	let altAlone = false;
	const tipsAt = (container: Element | null): HTMLElement[] =>
		deepTips(container ?? scope).filter((el) => {
			// A control inside a shared element's shadow root belongs where its host is.
			const owner = (hostOf.get(el) ?? el.parentElement)?.closest('[data-keytip-level]') ?? null;
			return owner === container && visible(el);
		});
	const current = () => tipsAt(levels.at(-1) ?? null);
	const render = () => {
		if (!handle.active) return;
		layer.replaceChildren(
			...current().map((el) => {
				const box = el.getBoundingClientRect();
				const tip = doc.createElement('span');
				tip.textContent = el.dataset.keytip!;
				tip.style.cssText = `${TIP_CSS};left:${Math.round(box.left + box.width / 2)}px;top:${Math.round(box.bottom - 4)}px;opacity:${isDisabled(el) || (typed && !el.dataset.keytip!.startsWith(typed)) ? '0.45' : '1'}`;
				return tip;
			}),
		);
	};
	const show = () => {
		if (topLayer && !layer.matches(':popover-open')) layer.showPopover?.();
		else layer.hidden = false;
		render();
	};
	const stop = () => {
		levels = [];
		typed = '';
		if (topLayer) {
			try {
				layer.hidePopover?.();
			} catch {
				/* Already hidden. */
			}
		} else layer.hidden = true;
		layer.replaceChildren();
		handle.active = false;
	};
	const start = () => {
		levels = [];
		typed = '';
		handle.active = true;
		show();
	};
	const choose = (el: HTMLElement) => {
		typed = '';
		if (isDisabled(el)) return render();
		const panel = el.dataset.keytipPanel;
		activate(el);
		const next = panel
			? ([...scope.querySelectorAll('[data-keytip-level]')].find((el) => el.id === panel) ?? null)
			: null;
		if (next) {
			levels.push(next);
			// Let the newly shown panel lay out before measuring its controls.
			view?.requestAnimationFrame ? view.requestAnimationFrame(render) : render();
		} else stop();
	};
	const onKeyDown = (event: Event) => {
		const key = event as KeyboardEvent;
		if (key.key === 'Alt') {
			altAlone = !key.ctrlKey && !key.metaKey && !key.shiftKey;
			return;
		}
		altAlone = false;
		if (!handle.active) return;
		if (key.key === 'Escape') {
			key.preventDefault();
			key.stopPropagation();
			typed = '';
			if (levels.length) {
				levels.pop();
				render();
			} else stop();
			return;
		}
		if (key.key.length !== 1 || key.ctrlKey || key.metaKey) return stop();
		key.preventDefault();
		key.stopPropagation();
		typed += key.key.toUpperCase();
		const candidates = current().filter((el) => el.dataset.keytip!.toUpperCase().startsWith(typed));
		const exact = candidates.find((el) => el.dataset.keytip!.toUpperCase() === typed);
		if (exact) choose(exact);
		else if (!candidates.length) typed = '';
		render();
	};
	const onKeyUp = (event: Event) => {
		const key = event as KeyboardEvent;
		if (key.key !== 'Alt' || !altAlone) return;
		altAlone = false;
		key.preventDefault();
		if (handle.active) stop();
		else start();
	};
	const onPointer = () => {
		if (handle.active) stop();
	};
	const target = scope as unknown as EventTarget;
	target.addEventListener('keydown', onKeyDown, true);
	target.addEventListener('keyup', onKeyUp, true);
	target.addEventListener('pointerdown', onPointer, true);
	view?.addEventListener('blur', onPointer);
	const handle = {
		active: false,
		start() {
			if (!handle.active) start();
		},
		stop() {
			stop();
		},
		dispose() {
			stop();
			target.removeEventListener('keydown', onKeyDown, true);
			target.removeEventListener('keyup', onKeyUp, true);
			target.removeEventListener('pointerdown', onPointer, true);
			view?.removeEventListener('blur', onPointer);
			layer.remove();
		},
	};
	return handle;
}
