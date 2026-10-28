/** A modal command's lifecycle, independent of its fields, translations and document model. */
export interface DialogSessionOptions<T> {
	element: HTMLElement;
	confirm?: HTMLButtonElement | undefined;
	cancel?: HTMLElement | undefined;
	submit?(): T | undefined | PromiseLike<T | undefined>;
	/** Product cleanup and focus restoration run once, including dismissal during an async submit. */
	finish(result: T | undefined): void;
	error?(error: unknown): void;
	/** Word permits Enter from radio choices; other dialogs normally confirm only text fields. */
	enter?(event: KeyboardEvent): boolean;
}

export function createDialogSession<T>(options: DialogSessionOptions<T>) {
	const { element, confirm, cancel } = options;
	let done = false;
	let submitting = false;
	let resolve!: (result: T | undefined) => void;
	const result = new Promise<T | undefined>((settle) => {
		resolve = settle;
	});
	const listeners: Array<() => void> = [];
	const listen = (target: EventTarget, name: string, handler: EventListener): void => {
		target.addEventListener(name, handler);
		listeners.push(() => target.removeEventListener(name, handler));
	};
	const close = (value?: T): void => {
		if (done) return;
		done = true;
		for (const remove of listeners) remove();
		options.finish(value);
		resolve(value);
	};
	const settle = (value: T | undefined): void => {
		if (!done && value !== undefined) close(value);
	};
	const failed = (error: unknown): void => {
		if (!done) options.error?.(error);
	};
	const submit = (): void => {
		if (done || submitting) return;
		if (!options.submit) return close();
		try {
			const outcome = options.submit();
			if (outcome && typeof (outcome as PromiseLike<T>).then === 'function') {
				submitting = true;
				const disabled = confirm?.disabled;
				if (confirm) confirm.disabled = true;
				void Promise.resolve(outcome)
					.then(settle, failed)
					.finally(() => {
						submitting = false;
						if (confirm && !done) confirm.disabled = disabled ?? false;
					});
			} else settle(outcome as T | undefined);
		} catch (error) {
			failed(error);
		}
	};
	listen(element, 'office-dialog-close', (event) => {
		event.preventDefault();
		close();
	});
	if (cancel) listen(cancel, 'click', () => close());
	if (confirm) listen(confirm, 'click', submit);
	listen(element, 'keydown', (raw) => {
		const event = raw as KeyboardEvent;
		if (event.defaultPrevented) return;
		if (event.key === 'Escape') {
			event.preventDefault();
			close();
		} else if (event.key === 'Enter' && confirm && !confirm.disabled) {
			const target = event.target as HTMLInputElement | null;
			const enter =
				options.enter?.(event) ??
				(target?.tagName === 'INPUT' && !['checkbox', 'radio', 'file'].includes(target.type));
			if (enter) {
				event.preventDefault();
				submit();
			}
		}
	});
	return { result, submit, close, dispose: () => close() };
}
