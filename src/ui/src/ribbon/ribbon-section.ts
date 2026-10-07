import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base';
import { defineButton } from './button';
import { defineRibbonGroup } from './ribbon-group';
import { definer } from '../registry';
import { tok } from '../tokens';
import css from './ribbon-section.css?raw';

/** One ribbon command, translated. Booleans map to the command element's attributes. */
export interface OfficeRibbonCommandView {
	id: string;
	label: string;
	/** ScreenTip; defaults to `label`. */
	title?: string | undefined;
	icon: string;
	badge?: string | number | undefined;
	/** `large`, `small` or `icon`; omitted leaves the command element's default. */
	size?: string | undefined;
	disabled?: boolean | undefined;
	active?: boolean | undefined;
	compact?: boolean | undefined;
	hidden?: boolean | undefined;
	caret?: boolean | undefined;
	/** Toggle state; undefined means not a toggle. */
	pressed?: boolean | undefined;
	/** Menu state; undefined means no menu. */
	expanded?: boolean | undefined;
	/** Stack commands with the same column number vertically. */
	column?: number | undefined;
}

export interface OfficeRibbonGroupView {
	id: string;
	label: string;
	commands: readonly OfficeRibbonCommandView[];
}

/** The vertical stack for commands that share a column; light DOM, so tokens apply inline. */
const COLUMN_STYLE = { display: 'flex', flexDirection: 'column', gap: tok('--office-space-0') };

interface GroupNode {
	el: HTMLElement;
	columns: Map<number, HTMLElement>;
}
type Statics = {
	groupTag: string;
	commandTag: string;
	groupIdAttribute: string;
	commandIdAttribute: string;
};

function setOptional(el: HTMLElement, name: string, value: boolean | undefined): void {
	if (value === undefined) el.removeAttribute(name);
	else el.setAttribute(name, String(value));
}

function syncCommand(el: HTMLElement, command: OfficeRibbonCommandView, idAttribute: string): void {
	el.setAttribute(idAttribute, command.id);
	el.setAttribute('label', command.label);
	el.setAttribute('title', command.title ?? command.label);
	el.setAttribute('icon', command.icon);
	if (command.badge) el.setAttribute('badge', String(command.badge));
	else el.removeAttribute('badge');
	if (command.size) el.setAttribute('size', command.size);
	for (const name of ['disabled', 'active', 'compact', 'hidden', 'caret'] as const)
		el.toggleAttribute(name, command[name] === true);
	setOptional(el, 'pressed', command.pressed);
	setOptional(el, 'expanded', command.expanded);
}

/**
 * `<office-ui-ribbon-section>`: a run of ribbon groups from data. Set `groups`
 * (`OfficeRibbonGroupView[]`); groups and commands are keyed by id and patched in place, so a
 * focused command and every customization id survive updates. Commands with a `column` stack
 * vertically. Children are real `office-ui-ribbon-group` and `office-ui-button` elements in the
 * light DOM, so their events (`office-command` and so on) bubble as usual. Static `groupTag`,
 * `commandTag`, `groupIdAttribute` and `commandIdAttribute` let a product keep its aliases.
 */
export class OfficeUiRibbonSection extends OfficeElement {
	static override styles = controlStyles(css);
	static groupTag = 'office-ui-ribbon-group';
	static commandTag = 'office-ui-button';
	static groupIdAttribute = 'data-ribbon-group';
	static commandIdAttribute = 'command';
	#groups: readonly OfficeRibbonGroupView[] = [];
	readonly #groupNodes = new Map<string, GroupNode>();
	readonly #commandNodes = new Map<string, HTMLElement>();
	protected override render() {
		return html`<slot></slot>`;
	}
	get groups(): readonly OfficeRibbonGroupView[] {
		return this.#groups;
	}
	set groups(value: readonly OfficeRibbonGroupView[] | null | undefined) {
		this.#groups = value ?? [];
		this.#render();
	}
	override connectedCallback(): void {
		super.connectedCallback();
		this.#render();
	}
	/**
	 * The groups and commands are light-DOM children keyed by id and patched in place (so focus
	 * and a product's tags and id attributes survive), which a shadow template cannot express.
	 */
	#render(): void {
		const statics = this.constructor as unknown as Statics;
		const doc = this.ownerDocument;
		const model = this.#groups;
		const wantedGroups = new Set(model.map((group) => group.id));
		const wantedCommands = new Set(model.flatMap((g) => g.commands.map((c) => c.id)));
		for (const [id, el] of this.#commandNodes)
			if (!wantedCommands.has(id)) {
				el.remove();
				this.#commandNodes.delete(id);
			}
		for (const [id, group] of this.#groupNodes)
			if (!wantedGroups.has(id)) {
				group.el.remove();
				this.#groupNodes.delete(id);
			}
		for (const [index, group] of model.entries()) {
			let node = this.#groupNodes.get(group.id);
			if (!node) {
				node = { el: doc.createElement(statics.groupTag), columns: new Map() };
				this.#groupNodes.set(group.id, node);
			}
			node.el.setAttribute(statics.groupIdAttribute, group.id);
			node.el.setAttribute('label', group.label);
			if (this.children[index] !== node.el)
				this.insertBefore(node.el, this.children[index] ?? null);
			const offsets = new Map<HTMLElement, number>();
			for (const command of group.commands) {
				let parent: HTMLElement = node.el;
				if (command.column !== undefined) {
					let column = node.columns.get(command.column);
					if (!column) {
						column = doc.createElement('div');
						Object.assign(column.style, COLUMN_STYLE);
						node.columns.set(command.column, column);
						node.el.append(column);
					}
					parent = column;
				}
				let el = this.#commandNodes.get(command.id);
				if (!el) {
					el = doc.createElement(statics.commandTag);
					this.#commandNodes.set(command.id, el);
				}
				syncCommand(el, command, statics.commandIdAttribute);
				const offset = offsets.get(parent) ?? 0;
				if (parent.children[offset] !== el)
					parent.insertBefore(el, parent.children[offset] ?? null);
				offsets.set(parent, offset + 1);
			}
		}
	}
}

export const defineRibbonSection = definer(
	'office-ui-ribbon-section',
	() => OfficeUiRibbonSection,
	[defineRibbonGroup, defineButton],
);
