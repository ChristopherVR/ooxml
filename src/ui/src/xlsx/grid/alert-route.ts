// Where a refused cell edit's data-validation alert is shown. The full editor draws the built-in
// modal (`alert.ts`); a surface (building-block) editor registers a route that hands the alert to
// the host's `onOpenValidationAlert`, or, without one, answers Retry so nothing is drawn.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { showValidationAlert, type AlertAnswer, type AlertStyle } from './alert';
import type { ValidationAlertRequest } from '../surface-actions';

export interface AlertQuery {
	style: AlertStyle;
	message: string;
	title?: string | undefined;
	ref: string;
	sheet: number;
	sheetName: string;
	value: string;
}
type Opener = ((request: ValidationAlertRequest) => void) | undefined;

const routes = new WeakMap<EditorContext, () => Opener>();

/** Makes `ctx` a surface: its alerts go to the opener `current()` returns, never to the canvas. */
export function routeValidationAlerts(ctx: EditorContext, current: () => Opener): void {
	routes.set(ctx, current);
}

/**
 * Asks for the user's answer: a promise for the built-in alert, or an immediate answer when the
 * surface has no host opener (Retry) or the opener throws.
 */
export function requestValidationAlert(
	ctx: EditorContext,
	query: AlertQuery,
): AlertAnswer | Promise<AlertAnswer> {
	const route = routes.get(ctx);
	if (!route) return showValidationAlert(ctx, query.style, query.message, query.title);
	const open = route();
	if (!open) return 'retry';
	return new Promise<AlertAnswer>((resolve) => {
		let settled = false;
		const answer = (value: AlertAnswer) => () => {
			if (settled) return;
			settled = true;
			resolve(value);
		};
		const base = {
			ref: query.ref,
			sheet: query.sheet,
			sheetName: query.sheetName,
			value: query.value,
			title: query.title ?? '',
			message: query.message,
			cancel: answer('cancel'),
		};
		const request: ValidationAlertRequest =
			query.style === 'stop'
				? { ...base, style: 'stop', retry: answer('retry') }
				: query.style === 'warning'
					? { ...base, style: 'warning', accept: answer('yes'), retry: answer('no') }
					: { ...base, style: 'information', accept: answer('ok') };
		try {
			open(request);
		} catch {
			answer('retry')();
		}
	});
}
