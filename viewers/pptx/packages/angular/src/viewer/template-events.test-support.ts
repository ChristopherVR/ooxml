/**
 * template-events.test-support.ts: fire a DOM event at one tag of an Angular
 * component's template the way the compiled template would.
 *
 * This package has no TestBed (see `vitest.config.ts`), so a test cannot render
 * a template and dispatch at its inputs. A source-text guard says a binding is
 * there; this helper goes one step further and RUNS it: it reads the
 * `(event)="handler(args)"` bindings off the tag and calls them on a component
 * instance, so a test can say "an `input` event on the fill picker moves the
 * fill" and fail, with the actual colour, when the template binds no `input`.
 *
 * Handler expressions are the shapes this package's templates use: a method
 * call whose arguments are `$event`, a literal, or a dotted path read from the
 * template scope (`section.id`) or the component (`fillColorInput`).
 *
 * Not a `.test.ts` file on purpose: a shared helper is not a suite.
 */

/** The `(event)` bindings of the first tag in `source` containing `marker`, by event name. */
export function templateEventBindings(source: string, marker: string): Map<string, string> {
	const at = source.indexOf(marker);
	if (at < 0) {
		throw new Error(`no template tag contains ${marker}`);
	}
	const start = source.lastIndexOf('<', at);
	const end = source.indexOf('>', at);
	const tag = source.slice(start, end);
	const bindings = new Map<string, string>();
	for (const match of tag.matchAll(/\((\w+)\)="([^"]*)"/gu)) {
		bindings.set(match[1]!, match[2]!);
	}
	return bindings;
}

function resolvePath(path: string, roots: readonly object[]): unknown {
	const [head = '', ...rest] = path.split('.');
	const root = roots.find((candidate) => head in candidate) as Record<string, unknown> | undefined;
	if (!root) {
		throw new Error(`template expression reads unknown name ${head}`);
	}
	return rest.reduce<unknown>(
		(value, key) => (value as Record<string, unknown> | undefined)?.[key],
		root[head],
	);
}

function argumentValue(arg: string, $event: unknown, roots: readonly object[]): unknown {
	if (arg === '$event') {
		return $event;
	}
	if (/^'.*'$/u.test(arg)) {
		return arg.slice(1, -1);
	}
	if (/^-?\d+(?:\.\d+)?$/u.test(arg)) {
		return Number(arg);
	}
	return resolvePath(arg, roots);
}

/**
 * Run the binding for the `type` event (a DOM event or a child's output) from
 * `bindings` on `component`, with `$event` as the template's `$event`. Does
 * nothing when the tag binds no such event, which is the point: the test then
 * observes that nothing happened.
 */
export function fireTemplateEvent(
	component: object,
	bindings: Map<string, string>,
	type: string,
	$event: unknown,
	scope: Record<string, unknown> = {},
): void {
	const expression = bindings.get(type);
	if (!expression) {
		return;
	}
	const call = /^\s*([\w.]+)\((.*)\)\s*$/u.exec(expression);
	if (!call) {
		throw new Error(`unsupported template expression ${expression}`);
	}
	const roots = [scope, component];
	const handler = resolvePath(call[1]!, roots);
	if (typeof handler !== 'function') {
		throw new Error(`${call[1]} is not a method`);
	}
	const args = call[2]!.trim()
		? call[2]!.split(',').map((arg) => argumentValue(arg.trim(), $event, roots))
		: [];
	(handler as (...values: unknown[]) => void).apply(component, args);
}

/** A DOM event of `type` whose target is a colour input showing `value`. */
export function colorInputEvent(type: 'input' | 'change' | 'blur', value: string): Event {
	const input = document.createElement('input');
	input.type = 'color';
	input.value = value;
	const event = new Event(type);
	Object.defineProperty(event, 'target', { value: input });
	return event;
}
