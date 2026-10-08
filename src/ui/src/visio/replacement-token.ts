import {
	VISIO_TEXT_REPLACE_LIMITS,
	visioTextReplaceCommands,
	visioTextReplaceOccurrences,
	visioTextReplacePlan,
	type VisioTextOccurrence,
	type VisioTextReplacePlan,
	type VisioTextReplaceScope,
} from 'ooxml-core/visio/ui';
import type { VisioDocument, VisioEdit } from 'ooxml-core/visio';

const replacementToken: unique symbol = Symbol('Visio replacement token');
/** Internal owned scope; contains no retained source bytes or document model. */
export interface ViewerTextReplaceToken {
	readonly [replacementToken]: true;
}
export type ViewerTextReplaceScope = VisioTextReplaceScope['scope'];
export interface ReplacementContext {
	readonly sourceGeneration: number;
	readonly documentGeneration: number;
	readonly operationGeneration: number;
	readonly selectionIntent: number;
	readonly navigationIntent: number;
	readonly pageId: string;
	readonly pageIndex: number;
}
interface Scope {
	readonly scope: ViewerTextReplaceScope;
	readonly pageId: string;
	readonly selection: readonly { readonly id: string; readonly pageId?: string }[];
}
interface Record {
	readonly context: ReplacementContext;
	readonly scope: Scope;
}
export interface ViewerTextReplacementInput {
	readonly query: string;
	readonly replacement: string;
	readonly mode: 'current' | 'all';
	readonly current?: VisioTextOccurrence;
}
interface Bridge {
	context(): ReplacementContext | undefined;
	document(): VisioDocument | null;
	selection(): Scope['selection'];
	navigate(occurrence: VisioTextOccurrence, context: ReplacementContext): ReplacementContext;
	mutate(
		edits: readonly VisioEdit[],
		next: VisioTextOccurrence | undefined,
		context: ReplacementContext,
		owner: object,
	): Promise<ReplacementContext>;
	cancel(owner: object): void;
}
export const replacementCancelled = () =>
	new DOMException('The text replacement was superseded or cancelled.', 'AbortError');
export function sameReplacementContext(a: ReplacementContext, b: ReplacementContext): boolean {
	return (
		a.sourceGeneration === b.sourceGeneration &&
		a.documentGeneration === b.documentGeneration &&
		a.operationGeneration === b.operationGeneration &&
		a.selectionIntent === b.selectionIntent &&
		a.navigationIntent === b.navigationIntent &&
		a.pageId === b.pageId &&
		a.pageIndex === b.pageIndex
	);
}

/** Controller orchestration only: core owns scopes, occurrence matching and command planning. */
export class ViewerTextReplacement {
	readonly #tokens = new WeakMap<ViewerTextReplaceToken, Record>();
	readonly #plans = new WeakMap<VisioTextReplacePlan, Record>();
	readonly #operations = new WeakMap<ViewerTextReplaceToken, object>();
	constructor(private readonly bridge: Bridge) {}
	capture(scope: ViewerTextReplaceScope): ViewerTextReplaceToken {
		const context = this.bridge.context();
		if (!context) throw replacementCancelled();
		if (!['selection', 'current-page', 'all-pages'].includes(scope))
			throw new Error('Choose a supported text replacement scope.');
		const selection = scope === 'selection' ? this.bridge.selection() : [];
		if (selection.length > VISIO_TEXT_REPLACE_LIMITS.shapes)
			throw new Error('Replace selection exceeds supported limits.');
		const owned = Object.freeze(
			selection.map(({ id, pageId }) =>
				Object.freeze({ id, ...(pageId === undefined ? {} : { pageId }) }),
			),
		);
		this.#requireContext(context);
		return this.#issue(context, Object.freeze({ scope, pageId: context.pageId, selection: owned }));
	}
	current(token: ViewerTextReplaceToken): boolean {
		const saved = this.#tokens.get(token);
		const context = saved && this.bridge.context();
		return !!saved && !!context && sameReplacementContext(saved.context, context);
	}
	/** Cancellation is bound to this token's application, even after scope intent changes. */
	cancel(token: ViewerTextReplaceToken): void {
		const owner = this.#operations.get(token);
		if (owner) this.bridge.cancel(owner);
	}
	occurrences(token: ViewerTextReplaceToken, query: string): readonly VisioTextOccurrence[] {
		const record = this.#require(token);
		return this.#read(token, () =>
			visioTextReplaceOccurrences(this.bridge.document()!, this.#scope(record, query)),
		);
	}
	plan(token: ViewerTextReplaceToken, input: ViewerTextReplacementInput): VisioTextReplacePlan {
		const record = this.#require(token);
		const plan = this.#read(token, () =>
			visioTextReplacePlan(this.bridge.document()!, {
				...this.#scope(record, input.query),
				replacement: input.replacement,
				mode: input.mode,
				...(input.current === undefined ? {} : { current: input.current }),
			}),
		);
		this.#plans.set(plan, record);
		return plan;
	}
	navigate(
		token: ViewerTextReplaceToken,
		occurrence: VisioTextOccurrence,
		query: string,
	): ViewerTextReplaceToken {
		const record = this.#require(token);
		const found = this.occurrences(token, query).find(
			(match) =>
				match.pageId === occurrence.pageId &&
				match.shapeId === occurrence.shapeId &&
				match.start === occurrence.start &&
				match.end === occurrence.end,
		);
		this.#require(token);
		if (!found) throw new Error('The replacement occurrence is outside the captured scope.');
		const context = this.bridge.navigate(found, record.context);
		this.#requireContext(context);
		return this.#issue(context, record.scope);
	}
	async apply(
		plan: VisioTextReplacePlan,
		token: ViewerTextReplaceToken,
	): Promise<ViewerTextReplaceToken> {
		const record = this.#require(token);
		const planned = this.#plans.get(plan);
		if (
			!planned ||
			planned.scope !== record.scope ||
			!sameReplacementContext(planned.context, record.context)
		)
			throw new Error('Create a replacement plan for the current owned search context.');
		const commands = this.#read(token, () =>
			visioTextReplaceCommands(this.bridge.document()!, plan),
		);
		if (!commands.length) return this.#issue(record.context, record.scope);
		const owner = Object.freeze({});
		this.#operations.set(token, owner);
		try {
			const context = await this.bridge.mutate(
				commands,
				plan.nextOccurrence,
				record.context,
				owner,
			);
			this.#requireContext(context);
			return this.#issue(context, record.scope);
		} finally {
			this.#operations.delete(token);
		}
	}
	#scope(record: Record, query: string): VisioTextReplaceScope {
		return { ...record.scope, query, matchCase: true };
	}
	#issue(context: ReplacementContext, scope: Scope): ViewerTextReplaceToken {
		const token = Object.freeze({ [replacementToken]: true as const });
		this.#tokens.set(token, { context, scope });
		return token;
	}
	#requireContext(saved: ReplacementContext): void {
		const current = this.bridge.context();
		if (!current || !sameReplacementContext(saved, current)) throw replacementCancelled();
	}
	#require(token: ViewerTextReplaceToken): Record {
		const record = this.#tokens.get(token);
		if (!record) throw replacementCancelled();
		this.#requireContext(record.context);
		return record;
	}
	#read<T>(token: ViewerTextReplaceToken, read: () => T): T {
		try {
			const result = read();
			this.#require(token);
			return result;
		} catch (error) {
			this.#require(token);
			throw error;
		}
	}
}
