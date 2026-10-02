// Insert > Link: a web page or file address, a place in this workbook, or an e-mail address.
// Only safe schemes are accepted (http, https, mailto, ftp, file and scheme-less paths).
import {
	type Hyperlink,
	getCell,
	parseAddress,
	parseRange,
	rangesIntersect,
} from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { field, invalid, panel, radios, select, textInput } from './fields.js';
import { button, showDialog } from './frame.js';

const SAFE_SCHEMES = ['http:', 'https:', 'mailto:', 'ftp:', 'file:'];

/** The address to store, or undefined when its scheme is not allowed (javascript:, data:, ...). */
export function safeAddress(text: string): string | undefined {
	const value = text.trim();
	if (!value) return undefined;
	if (/^www\./i.test(value)) return `https://${value}`;
	const scheme = /^([a-z][a-z0-9+.-]*):/i.exec(value)?.[1];
	// A single letter is a Windows drive (C:\...), not a scheme.
	if (!scheme || scheme.length === 1) return value;
	return SAFE_SCHEMES.includes(`${scheme.toLowerCase()}:`) ? value : undefined;
}

/** Splits `Sheet!A1` / `'My sheet'!A1` / a defined name into its sheet and reference parts. */
function splitLocation(location: string): { sheet?: string; ref: string } {
	const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(.*)$/.exec(location);
	if (!match) return { ref: location };
	return { sheet: (match[1] ?? match[2] ?? '').replace(/''/g, "'"), ref: match[3] ?? '' };
}

const quoteSheet = (name: string): string =>
	/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;

export function hyperlinkDialog(
	ctx: EditorContext,
): Promise<Omit<Hyperlink, 'range'> | null | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const existing = t.ws.hyperlinks.find((h) =>
		rangesIntersect(h.range, { start: t.active, end: t.active }),
	);
	const cell = getCell(t.ws, t.active.row, t.active.col);
	const cellText = cell?.value === null || cell?.value === undefined ? '' : String(cell.value);
	const mail = existing?.target?.toLowerCase().startsWith('mailto:') ? existing.target : undefined;
	const kind = existing?.location ? 'place' : mail ? 'email' : 'web';
	const linkTo = radios(
		ctx,
		'Link to:',
		[
			['web', 'Existing File or Web Page'],
			['place', 'Place in This Document'],
			['email', 'E-mail Address'],
		],
		kind,
	);
	const display = textInput(ctx, existing?.display ?? cellText);
	const tooltip = textInput(ctx, existing?.tooltip ?? '');
	const address = textInput(ctx, !mail ? (existing?.target ?? '') : '');
	const sheets = t.workbook.sheets.map((s) => [s.name, s.name] as const);
	const names = t.workbook.definedNames
		.filter((n) => !n.hidden && !n.name.startsWith('_xlnm.'))
		.map((n) => [`name:${n.name}`, n.name] as const);
	const loc = existing?.location
		? splitLocation(existing.location)
		: { sheet: t.ws.name, ref: 'A1' };
	const isName = !loc.sheet && names.some(([, n]) => n === loc.ref);
	const place = select(
		ctx,
		[...sheets, ...names],
		isName ? `name:${loc.ref}` : (loc.sheet ?? t.ws.name),
		true,
	);
	const ref = textInput(ctx, isName ? '' : loc.ref);
	const parsedMail = mail ? /^mailto:([^?]*)(?:\?subject=(.*))?$/i.exec(mail) : null;
	const email = textInput(ctx, parsedMail ? decodeURIComponent(parsedMail[1] ?? '') : '');
	const subject = textInput(ctx, parsedMail?.[2] ? decodeURIComponent(parsedMail[2]) : '');
	const pages = {
		web: panel(ctx, field(ctx, 'Address:', address)),
		place: panel(
			ctx,
			field(ctx, 'Type the cell reference:', ref),
			field(ctx, 'Or select a place in this document:', place),
		),
		email: panel(ctx, field(ctx, 'E-mail address:', email), field(ctx, 'Subject:', subject)),
	};
	const sync = (): void => {
		const v = linkTo.get();
		for (const [k, p] of Object.entries(pages)) p.hidden = k !== v;
		ref.disabled = place.value.startsWith('name:');
	};
	for (const input of linkTo.inputs) input.addEventListener('change', sync);
	place.addEventListener('change', sync);
	sync();
	let removed = false;
	const remove = button(ctx, 'Remove Link');
	const buttons = existing ? [remove] : [];
	const read = (): Omit<Hyperlink, 'range'> | undefined => {
		const link: Omit<Hyperlink, 'range'> = {};
		const v = linkTo.get();
		if (v === 'web') {
			if (!address.value.trim()) return invalid(ctx, address, 'Enter an address.');
			const safe = safeAddress(address.value);
			if (!safe) return invalid(ctx, address, 'This address uses a protocol that is not allowed.');
			link.target = safe;
		} else if (v === 'email') {
			const to = email.value.trim().replace(/^mailto:/i, '');
			if (!to || /[\s<>"]/.test(to)) return invalid(ctx, email, 'Enter a valid e-mail address.');
			link.target = `mailto:${to}${subject.value.trim() ? `?subject=${encodeURIComponent(subject.value.trim())}` : ''}`;
		} else if (place.value.startsWith('name:')) link.location = place.value.slice(5);
		else {
			const r = ref.value.trim().replace(/\$/g, '') || 'A1';
			if (!parseAddress(r) && !parseRange(r))
				return invalid(ctx, ref, 'The reference is not valid.');
			link.location = `${quoteSheet(place.value)}!${r.toUpperCase()}`;
		}
		const shown = display.value.trim();
		link.display = shown || link.target || link.location || '';
		if (tooltip.value.trim()) link.tooltip = tooltip.value.trim();
		return link;
	};
	return showDialog<Omit<Hyperlink, 'range'> | null>(
		ctx,
		{
			name: 'hyperlink',
			heading: existing ? 'Edit Hyperlink' : 'Insert Hyperlink',
			wide: true,
			body: [
				linkTo.element,
				field(ctx, 'Text to display:', display),
				pages.web,
				pages.place,
				pages.email,
				field(ctx, 'ScreenTip:', tooltip),
			],
			buttons,
			opened: () => (kind === 'web' ? address : kind === 'email' ? email : ref).focus(),
			submit: () => {
				const link = read();
				if (!link) return undefined;
				t.session.batch('Insert link', () => {
					if (shownChanged(cellText, display.value) && link.display)
						t.session.setCellValue(t.sheet, t.active.row, t.active.col, link.display);
					t.session.setHyperlink(t.sheet, existing?.range ?? t.range, link);
				});
				return link;
			},
		},
		(open) => {
			remove.addEventListener('click', () => {
				if (removed) return;
				removed = true;
				t.session.setHyperlink(t.sheet, existing?.range ?? t.range, undefined);
				open.close();
			});
		},
	);
}

const shownChanged = (cellText: string, display: string): boolean =>
	cellText !== '' && display.trim() !== '' && display.trim() !== cellText;
