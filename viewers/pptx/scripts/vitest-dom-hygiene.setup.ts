// Runs per test file on a shared worker (`isolate: false`): the DOM is reused across files, so
// whatever one file leaves on `document` or `window` (listeners, mounted nodes) would be seen by
// the next. Track the listeners still registered by this file and remove them when it ends, and
// clear the body, so a file starts from the DOM an isolated worker would have given it.
//
// Only live registrations are kept (removal and `once` drop the entry), so a test that mounts and
// destroys many viewers does not have its listeners, and everything they close over, retained.
import { afterAll } from 'vitest';

type Entry = {
	target: EventTarget;
	type: string;
	listener: EventListenerOrEventListenerObject;
	capture: boolean;
};

const captureOf = (options?: boolean | EventListenerOptions): boolean =>
	typeof options === 'boolean' ? options : Boolean(options?.capture);

if (typeof document !== 'undefined' && typeof window !== 'undefined') {
	// Keyed by listener so a lookup stays cheap however many listeners a file registers.
	const live = new Map<EventListenerOrEventListenerObject, Entry[]>();
	const matches = (entry: Entry, target: EventTarget, type: string, capture: boolean) =>
		entry.target === target && entry.type === type && entry.capture === capture;
	const restore: (() => void)[] = [];
	for (const target of [document, window] as EventTarget[]) {
		const add = target.addEventListener;
		const remove = target.removeEventListener;
		target.addEventListener = function (
			this: EventTarget,
			type: string,
			listener: EventListenerOrEventListenerObject | null,
			options?: boolean | AddEventListenerOptions,
		) {
			const once = typeof options === 'object' && options !== null && options.once === true;
			if (listener && !once) {
				const capture = captureOf(options);
				const entries = live.get(listener) ?? [];
				if (!entries.some((entry) => matches(entry, target, type, capture))) {
					entries.push({ target, type, listener, capture });
					live.set(listener, entries);
				}
			}
			return add.call(this, type, listener, options);
		} as EventTarget['addEventListener'];
		target.removeEventListener = function (
			this: EventTarget,
			type: string,
			listener: EventListenerOrEventListenerObject | null,
			options?: boolean | EventListenerOptions,
		) {
			const capture = captureOf(options);
			const entries = listener ? live.get(listener) : undefined;
			if (entries) {
				const kept = entries.filter((entry) => !matches(entry, target, type, capture));
				if (kept.length > 0) {
					live.set(listener!, kept);
				} else {
					live.delete(listener!);
				}
			}
			return remove.call(this, type, listener, options);
		} as EventTarget['removeEventListener'];
		restore.push(() => {
			target.addEventListener = add;
			target.removeEventListener = remove;
		});
	}
	afterAll(() => {
		for (const undo of restore) {
			undo();
		}
		for (const entries of live.values()) {
			for (const { target, type, listener, capture } of entries) {
				target.removeEventListener(type, listener, capture);
			}
		}
		live.clear();
		document.body?.replaceChildren();
	});
}
