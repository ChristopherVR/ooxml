import { afterEach, expect, it, vi } from 'vitest';
import { createDialogSession } from './dialog-session';

afterEach(() => document.body.replaceChildren());

it('serializes cross-realm thenable submissions and ignores late completion after dismissal', async () => {
	const element = document.createElement('div');
	const confirm = document.createElement('button');
	element.append(confirm);
	document.body.append(element);
	let complete!: (value: boolean) => void;
	const submit = vi.fn(
		() =>
			({
				then: (resolve: (value: boolean) => void) => {
					complete = resolve;
				},
			}) as PromiseLike<boolean>,
	);
	const finish = vi.fn();
	const session = createDialogSession({ element, confirm, submit, finish });
	session.submit();
	session.submit();
	expect(submit).toHaveBeenCalledOnce();
	expect(confirm.disabled).toBe(true);
	await Promise.resolve();
	session.dispose();
	expect(await session.result).toBeUndefined();
	complete(true);
	await Promise.resolve();
	await Promise.resolve();
	expect(finish).toHaveBeenCalledOnce();
	expect(finish).toHaveBeenCalledWith(undefined);
	confirm.click();
	expect(submit).toHaveBeenCalledOnce();
});

it('keeps invalid submissions open and routes Enter using the element owner document', async () => {
	const doc = document.implementation.createHTMLDocument();
	const element = doc.createElement('div');
	const confirm = doc.createElement('button');
	const input = doc.createElement('input');
	element.append(input, confirm);
	const submit = vi
		.fn<() => string | undefined>()
		.mockReturnValueOnce(undefined)
		.mockReturnValue('done');
	const finish = vi.fn();
	const session = createDialogSession({ element, confirm, submit, finish });
	input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(finish).not.toHaveBeenCalled();
	input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
	expect(await session.result).toBe('done');
	expect(finish).toHaveBeenCalledOnce();
});
