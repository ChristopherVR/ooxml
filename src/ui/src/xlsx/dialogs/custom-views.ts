// View > Custom Views: the list of saved views with Show, Add and Delete (as Excel's dialog).
import { customViewsOf, type CustomWorkbookView } from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, el, field, listBox, textInput } from './fields';
import { button, showDialog } from './frame';

type Target = NonNullable<ReturnType<typeof target>>;

const report = (ctx: EditorContext, error: unknown): void =>
	ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'warning');

/** The Add View dialog: a name and what the view includes. */
async function addView(ctx: EditorContext, t: Target): Promise<boolean> {
	const name = textInput(ctx, '');
	const print = checkbox(ctx, 'Print settings', true);
	const hidden = checkbox(ctx, 'Hidden rows, columns and filter settings', true);
	const submitted = await showDialog<boolean>(ctx, {
		name: 'custom-view-add',
		heading: 'Add View',
		body: [field(ctx, 'Name:', name), print.wrapper, hidden.wrapper],
		opened: () => name.focus(),
		submit: () => {
			try {
				t.session.addCustomView({
					name: name.value,
					includePrintSettings: print.input.checked,
					includeHiddenRowCol: hidden.input.checked,
				});
				return true;
			} catch (error) {
				report(ctx, error);
				return undefined;
			}
		},
	});
	return submitted === true;
}

export function openCustomViews(ctx: EditorContext): Promise<undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const list = listBox(
		ctx,
		'Views:',
		() => sync(),
		() => void show(),
	);
	const note = el(ctx, 'div', 'xve-note');
	const add = button(ctx, 'Add...');
	const del = button(ctx, 'Delete');
	const views = (): CustomWorkbookView[] => customViewsOf(t.workbook);
	const chosen = (): CustomWorkbookView | undefined => {
		const guid = list.value();
		return views().find((v) => v.guid === guid);
	};
	function sync(): void {
		del.disabled = !chosen();
		note.textContent = views().length ? '' : ctx.t('No custom views are saved.');
	}
	const refresh = (): void => {
		list.setItems(views().map((v) => [v.guid, v.name] as const));
		sync();
	};
	/** Shows the chosen view; false (with a message) when none is chosen or Excel would refuse. */
	const show = (): boolean => {
		const view = chosen();
		if (!view) {
			ctx.toast(ctx.t('Select a view to show.'), 'warning');
			return false;
		}
		try {
			t.session.showCustomView(view.guid);
			return true;
		} catch (error) {
			report(ctx, error);
			return false;
		}
	};
	add.addEventListener('click', () => void addView(ctx, t).then(refresh));
	del.addEventListener('click', () => {
		const view = chosen();
		if (!view) return;
		try {
			t.session.deleteCustomView(view.guid);
		} catch (error) {
			report(ctx, error);
		}
		refresh();
	});
	refresh();
	return showDialog<boolean>(ctx, {
		name: 'custom-views',
		heading: 'Custom Views',
		okLabel: 'Show',
		cancelLabel: 'Close',
		buttons: [add, del],
		body: [list.element, note],
		opened: () => list.element.focus(),
		// Show closes the dialog, as in Excel; a missing choice keeps it open.
		submit: () => (show() ? true : undefined),
	}).then(() => undefined);
}
