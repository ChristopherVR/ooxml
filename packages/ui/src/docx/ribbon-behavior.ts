import { showCommandTips } from './ribbon-keytips';
import { fileButton, panelOf, tabButtons } from './ribbon-tab-api';

/** Word's tab KeyTips (Alt, then a letter), by tab key. File is `F`. */
const KEY_TIPS: Record<string, string> = {
	home: 'H',
	insert: 'N',
	layout: 'P',
	references: 'S',
	review: 'R',
	view: 'W',
	table: 'T',
	'header-footer': 'J',
};

/** The badge style. The tabs live in the shared ribbon's shadow root, which editor CSS cannot reach. */
const BADGE_STYLE =
	'position:absolute;bottom:-2px;left:50%;z-index:70;transform:translateX(-50%);min-width:16px;' +
	'padding:1px 4px;border:1px solid var(--office-border,#ccc);border-radius:3px;background:#1f1f1f;' +
	"color:#fff;font:600 11px/14px 'Segoe UI',Arial,sans-serif;text-align:center;pointer-events:none";

/**
 * Ribbon KeyTips. The shared ribbon owns collapse and peek (`collapsible`); here, when the ribbon
 * is focused with Alt or F10 (the host fires `dve-keytips`), tab badges appear and the matching
 * letter opens that tab (`F` opens File); the tab's commands then show their own tips (see
 * ribbon-keytips.ts).
 */
export function attachRibbonBehavior(root: HTMLElement): void {
	root.toggleAttribute('collapsible', true);
	let hideTips: (() => void) | undefined;
	root.addEventListener('dve-keytips', () => {
		hideTips?.();
		const targets: Array<{ tab: HTMLElement; key: string | undefined; panel: string | undefined }> =
			[
				...(fileButton(root) ? [{ tab: fileButton(root)!, key: 'F', panel: undefined }] : []),
				...tabButtons(root).map((tab) => ({
					tab,
					key: KEY_TIPS[tab.dataset.tab ?? ''],
					panel: tab.dataset.tab,
				})),
			];
		const badges = targets.flatMap(({ tab, key, panel }) => {
			if (!key) return [];
			const badge = document.createElement('span');
			badge.className = 'dve-keytip';
			badge.style.cssText = BADGE_STYLE;
			badge.textContent = key;
			badge.setAttribute('aria-hidden', 'true');
			tab.append(badge);
			return [{ badge, key, tab, panel }];
		});
		const stop = () => {
			for (const { badge } of badges) badge.remove();
			root.removeEventListener('keydown', onKey, true);
			root.removeEventListener('pointerdown', stop, true);
			root.removeEventListener('focusout', onFocusOut);
			hideTips = undefined;
		};
		const onKey = (event: KeyboardEvent) => {
			if (event.key === 'Alt' || event.key === 'Shift') return;
			const match = badges.find(({ key }) => key === event.key.toUpperCase());
			stop();
			if (!match || event.ctrlKey || event.metaKey) return;
			event.preventDefault();
			event.stopPropagation();
			match.tab.click();
			match.tab.focus();
			// Word's second level: the tab's commands get their own tips.
			const panel = match.panel ? panelOf(root, match.panel) : null;
			if (panel) showCommandTips(root, panel, () => match.tab.focus());
		};
		const onFocusOut = (event: FocusEvent) => {
			if (!event.relatedTarget || !root.contains(event.relatedTarget as Node)) stop();
		};
		root.addEventListener('keydown', onKey, true);
		root.addEventListener('pointerdown', stop, true);
		root.addEventListener('focusout', onFocusOut);
		hideTips = stop;
	});
}
