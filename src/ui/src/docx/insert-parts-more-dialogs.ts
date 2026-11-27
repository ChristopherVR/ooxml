import {
	DOCUMENT_PROPERTY_CHOICES,
	docInfoFieldResult,
	docPropertyResult,
	resolveWebVideo,
} from 'ooxml-core/docx';
import { dialogButton, labelled, selectOf } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { formDialog, note, type InsertPartsHost } from './insert-parts-form';
import { createVideoSearchPanel } from './online-video-search-panel';
import {
	localizeElement,
	translate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';

/** The fields the Field dialog inserts: those this package computes. */
const COMPUTED_FIELDS = [
	['DOCPROPERTY', 'DocProperty'],
	['AUTHOR', 'Author'],
	['COMMENTS', 'Comments'],
	['KEYWORDS', 'Keywords'],
	['LASTSAVEDBY', 'LastSavedBy'],
	['SUBJECT', 'Subject'],
	['TITLE', 'Title'],
] as const;

/** Word fields this package does not compute: listed so the choice is visible, but not available. */
const UNAVAILABLE_FIELDS = [
	'Date',
	'Formula',
	'Hyperlink',
	'IncludeText',
	'MergeField',
	'NumPages',
	'Page',
	'Ref',
	'Seq',
	'StyleRef',
	'TOC',
] as const;

/** Insert > Quick Parts > Field... */
export function createFieldDialog(host: InsertPartsHost): FormatDialog {
	const names = document.createElement('select');
	names.size = 10;
	names.style.minHeight = '13em';
	for (const [value, label] of COMPUTED_FIELDS) names.append(new Option(label, value));
	for (const label of UNAVAILABLE_FIELDS) {
		const option = new Option(label, label);
		option.disabled = true;
		names.append(option);
	}
	const property = selectOf([]);
	const ok = dialogButton('OK', true);
	const cancel = dialogButton('Cancel');
	const propertyRow = labelled('Property', property);
	const sync = () => {
		propertyRow.hidden = names.value !== 'DOCPROPERTY';
	};
	names.addEventListener('change', sync);
	return formDialog(host, {
		title: 'Field',
		className: 'dve-field-dialog',
		content: [
			labelled('Field names', names),
			propertyRow,
			note('Fields shown in grey are not computed by this editor.'),
		],
		ok,
		cancel,
		first: names,
		inputs: [names, property],
		reset() {
			const custom = host.properties().custom.map((item) => item.name);
			const builtIn = DOCUMENT_PROPERTY_CHOICES.map((choice) => choice.label);
			property.replaceChildren(...[...builtIn, ...custom].map((label) => new Option(label, label)));
			names.selectedIndex = 0;
			sync();
			return true;
		},
		valid: () => COMPUTED_FIELDS.some(([value]) => value === names.value),
		apply() {
			const name = names.value;
			const instr =
				name === 'DOCPROPERTY' ? ` DOCPROPERTY ${quoted(property.value)} ` : ` ${name} `;
			const properties = host.properties();
			const result = docPropertyResult(instr, properties) ?? docInfoFieldResult(instr, properties);
			host.insertField(instr, result?.trim() ? result : ' ');
			return true;
		},
	});
}

const quoted = (name: string): string =>
	/^\S+$/.test(name) ? name : `"${name.replace(/"/g, '')}"`;

/** Insert > Online Video: an embed code or address, checked against the host's allow-list. */
export function createOnlineVideoDialog(host: InsertPartsHost): FormatDialog {
	const embed = document.createElement('textarea');
	embed.rows = 3;
	embed.spellcheck = false;
	const allowed = note('');
	const message = note('');
	message.setAttribute('role', 'alert');
	const ok = dialogButton('Insert', true);
	const cancel = dialogButton('Cancel');
	let locale: EditorLocale = 'en';
	const check = () => resolveWebVideo(embed.value, host.onlineVideoHosts());
	const search = createVideoSearchPanel(host.onlineVideoHosts, (result) => {
		embed.value = result.url;
		embed.dispatchEvent(new Event('input'));
		embed.focus();
	});
	const dialog = formDialog(host, {
		title: 'Insert Video',
		className: 'dve-online-video-dialog',
		content: [search.element, labelled('Video embed code or address', embed), allowed, message],
		ok,
		cancel,
		first: embed,
		inputs: [embed],
		reset() {
			embed.value = '';
			allowed.textContent = `${translate(locale, 'Allowed sites' as LocalizationKey)}: ${host.onlineVideoHosts().join(', ')}`;
			message.textContent = '';
			search.reset(host.onlineVideoSearch(), locale);
			return host.onlineVideoHosts().length > 0;
		},
		valid() {
			const result = embed.value.trim() ? check() : undefined;
			message.textContent =
				!result || result.ok
					? ''
					: result.reason === 'host'
						? 'This video site is not allowed by the host of this editor.'
						: 'Paste an embed code or an https address of the video player.';
			localizeElement(message, locale);
			return Boolean(result?.ok);
		},
		apply() {
			const result = check();
			if (!result.ok) return false;
			search.cancel();
			host.insertWebVideo(result.spec);
			return true;
		},
	});
	return {
		...dialog,
		setLocale(next) {
			locale = next;
			dialog.setLocale(next);
		},
		get isOpen() {
			return dialog.isOpen;
		},
	};
}
