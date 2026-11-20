/**
 * Limits on one assistant run: everything from a user message until the model
 * stops. Without them a model that keeps calling tools runs until the provider
 * or the bill stops it. Every connection kind goes through this one limiter.
 */

/** Per-run limits. Omitted fields use {@link DEFAULT_RUN_LIMITS}. */
export interface PptxAiRunLimits {
	/** Model requests in one run. */
	maxSteps?: number;
	/** Tool calls in one run. */
	maxToolCalls?: number;
	/** Wall-clock time for one run, in milliseconds. */
	maxRunMs?: number;
	/** How often one run may make the same call with the same input. */
	maxRepeatedToolCalls?: number;
}

export const DEFAULT_RUN_LIMITS: Required<PptxAiRunLimits> = {
	maxSteps: 16,
	maxToolCalls: 48,
	maxRunMs: 600_000,
	maxRepeatedToolCalls: 4,
};

/** Which limit ended a run. */
export type PptxAiRunLimit = 'steps' | 'tool-calls' | 'time' | 'repeats';

/** How the latest model request of a run ended, in one vocabulary for every provider. */
export type PptxAiFinishReason =
	| 'stop'
	| 'tool-calls'
	| 'length'
	| 'filtered'
	| 'cancelled'
	| 'error'
	| 'limit'
	| 'unknown';

export interface PptxAiRunFinish {
	reason: PptxAiFinishReason;
	/** Present when `reason` is `'limit'`. */
	limit?: PptxAiRunLimit;
}

const LIMIT_MESSAGES: Record<PptxAiRunLimit, string> = {
	steps: 'The assistant stopped because it reached its limit of steps for one request.',
	'tool-calls': 'The assistant stopped because it reached its limit of tool calls for one request.',
	time: 'The assistant stopped because this request took too long.',
	repeats: 'The assistant stopped because it kept repeating the same action.',
};

/** Thrown instead of running a tool once the run is over its limits. */
export class RunLimitError extends Error {
	constructor(readonly limit: PptxAiRunLimit) {
		super(LIMIT_MESSAGES[limit]);
		this.name = 'RunLimitError';
	}
}

export class RunLimiter {
	private readonly limits: Required<PptxAiRunLimits>;
	private steps = 0;
	private toolCalls = 0;
	private calls = new Map<string, number>();
	private deadline: number | undefined;
	private hit: PptxAiRunLimit | undefined;

	constructor(
		limits: PptxAiRunLimits = {},
		private readonly now: () => number = Date.now,
	) {
		this.limits = { ...DEFAULT_RUN_LIMITS, ...limits };
	}

	/** The limit that ended this run, if one did. */
	get tripped(): PptxAiRunLimit | undefined {
		return this.hit;
	}

	get maxSteps(): number {
		return this.limits.maxSteps;
	}

	/** Whether a run has begun. */
	get started(): boolean {
		return this.deadline !== undefined;
	}

	/** Start counting a new run. */
	begin(): void {
		this.steps = 0;
		this.toolCalls = 0;
		this.calls = new Map();
		this.hit = undefined;
		this.deadline = this.now() + this.limits.maxRunMs;
	}

	trip(limit: PptxAiRunLimit): void {
		this.hit ??= limit;
	}

	/** Milliseconds left in this run. */
	remainingMs(): number {
		return Math.max(0, (this.deadline ?? this.now()) - this.now());
	}

	private checkTime(): void {
		if (this.deadline !== undefined && this.now() >= this.deadline) {
			this.trip('time');
		}
	}

	/** Count a model request. Returns false when the run may not make another one. */
	admitStep(): boolean {
		this.checkTime();
		if (!this.hit && this.steps >= this.limits.maxSteps) {
			this.trip('steps');
		}
		if (this.hit) {
			return false;
		}
		this.steps += 1;
		return true;
	}

	/** Count a tool call, or throw {@link RunLimitError}. Calls outside a run are not limited. */
	admitToolCall(name: string, input: unknown): void {
		if (this.deadline === undefined) {
			return;
		}
		this.checkTime();
		this.toolCalls += 1;
		const key = `${name}:${JSON.stringify(input)}`;
		const repeats = (this.calls.get(key) ?? 0) + 1;
		this.calls.set(key, repeats);
		if (this.toolCalls > this.limits.maxToolCalls) {
			this.trip('tool-calls');
		} else if (repeats > this.limits.maxRepeatedToolCalls) {
			this.trip('repeats');
		}
		if (this.hit) {
			throw new RunLimitError(this.hit);
		}
	}
}
