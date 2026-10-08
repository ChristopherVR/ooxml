import { html } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer, present } from '../registry';
import css from './ribbon-actions.css?raw';

/** One choice of the editing-mode selector (Word's and Excel's Editing / Viewing). */
export interface OfficeRibbonActionsMode {
	value: string;
	label: string;
}

/** Emitted when the user picks another editing mode. */
export type OfficeRibbonModeEvent = CustomEvent<{ mode: string }>;

/**
 * `<office-ui-ribbon-actions>`: the controls Office places at the right end of the ribbon tab row,
 * in Office's order: the editing-mode selector (pencil icon and a native select), Comments (icon
 * and label) and the accent Share button. Put it in `office-ui-ribbon`'s `actions` slot (or at a
 * product tab row's right end). Controlled: the host owns every value and redraws after an event.
 *
 * Properties: `modes` (the selector's choices; empty hides it), `mode`, `mode-label` (its name),
 * `comments-label`, `comments-title`, `comments-pressed`, `comments-count` (a badge when above 0),
 * `no-comments`, `share-label`, `share-title`, `share-pressed` (a live session) and `no-share`.
 * Events: `office-ribbon-mode` `{ mode }`, `office-ribbon-comments` and `office-ribbon-share`.
 * Parts: `mode`, `mode-select`, `comments`, `share`. Below 900px the labels hide, as in Office's
 * narrow layout; the buttons keep their accessible names through `aria-label`.
 */
export class OfficeUiRibbonActions extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		modes: { attribute: false },
		mode: { type: String },
		modeLabel: { attribute: 'mode-label', type: String },
		commentsLabel: { attribute: 'comments-label', type: String },
		commentsTitle: { attribute: 'comments-title', type: String },
		commentsPressed: { attribute: 'comments-pressed', ...flag },
		commentsCount: { attribute: 'comments-count', type: Number },
		noComments: { attribute: 'no-comments', ...flag },
		shareLabel: { attribute: 'share-label', type: String },
		shareTitle: { attribute: 'share-title', type: String },
		sharePressed: { attribute: 'share-pressed', ...flag },
		noShare: { attribute: 'no-share', ...flag },
	};
	declare modes: OfficeRibbonActionsMode[];
	declare mode: string;
	declare modeLabel: string;
	declare commentsLabel: string;
	declare commentsTitle: string | null;
	declare commentsPressed: boolean;
	declare commentsCount: number;
	declare noComments: boolean;
	declare shareLabel: string;
	declare shareTitle: string | null;
	declare sharePressed: boolean;
	declare noShare: boolean;

	constructor() {
		super();
		this.modes = [];
		this.mode = '';
		this.modeLabel = 'Editing mode';
		this.commentsLabel = 'Comments';
		this.commentsTitle = null;
		this.commentsPressed = false;
		this.commentsCount = 0;
		this.noComments = false;
		this.shareLabel = 'Share';
		this.shareTitle = null;
		this.sharePressed = false;
		this.noShare = false;
	}

	/** The rendered control of a part, or null before the first render. */
	control(part: 'mode-select' | 'comments' | 'share'): HTMLElement | null {
		return this.renderRoot.querySelector<HTMLElement>(`[part="${part}"]`);
	}

	private onMode(event: Event): void {
		event.stopPropagation();
		const select = event.target as HTMLSelectElement;
		const mode = select.value;
		// Controlled: keep showing the host's mode until it redraws with the new one.
		select.value = this.mode;
		if (mode !== this.mode) this.fire('office-ribbon-mode', { mode });
	}

	protected override updated(): void {
		const select = this.control('mode-select') as HTMLSelectElement | null;
		if (select && select.value !== this.mode) select.value = this.mode;
	}

	protected override render() {
		const modes = this.modes ?? [];
		const count = Number(this.commentsCount) || 0;
		const commentsTitle = this.commentsTitle ?? this.commentsLabel;
		const shareTitle = this.shareTitle ?? this.shareLabel;
		// No whitespace between nodes: consumers compare the buttons' textContent.
		// prettier-ignore
		return html`<label class="mode" part="mode" ?hidden=${modes.length === 0}>${glyph('pencil', 'glyph')}<select class="mode-select" part="mode-select" aria-label=${this.modeLabel} title=${this.modeLabel} @change=${this.onMode}>${modes.map((option) => html`<option value=${option.value} ?selected=${option.value === this.mode}>${option.label}</option>`)}</select></label><button class="comments" part="comments" type="button" aria-label=${this.commentsLabel} title=${count > 0 ? `${commentsTitle} (${count})` : commentsTitle} aria-pressed=${String(present(this.commentsPressed))} ?hidden=${present(this.noComments)} @click=${() => this.fire('office-ribbon-comments', {})}>${glyph('message', 'glyph')}<span class="label">${this.commentsLabel}</span>${count > 0 ? html`<span class="badge" aria-hidden="true">${count}</span>` : ''}</button><button class="share" part="share" type="button" aria-label=${this.shareLabel} title=${shareTitle} aria-pressed=${String(present(this.sharePressed))} ?hidden=${present(this.noShare)} @click=${() => this.fire('office-ribbon-share', {})}>${glyph('share', 'glyph')}<span class="label">${this.shareLabel}</span></button>`;
	}
}

export const defineRibbonActions = definer('office-ui-ribbon-actions', () => OfficeUiRibbonActions);
