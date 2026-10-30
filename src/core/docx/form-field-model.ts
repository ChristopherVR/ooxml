// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The model of a legacy form field (`w:fldChar/w:ffData`): a Word 97 text box, check box or
// drop-down, held on the run with the field's begin marker. Anything not modeled stays in
// `sourceXml`.
import type { StDocProtect, StFFTextType, StInfoTextType } from './generated/wml-simple-types';

/** Help or status-bar text: the text itself, or the name of an AutoText entry holding it. */
export interface FormFieldInfoText {
	type: StInfoTextType;
	value: string;
}

export interface FormField {
	/** FORMTEXT, FORMCHECKBOX or FORMDROPDOWN. */
	kind: 'text' | 'checkbox' | 'dropDown';
	/** `w:name`, also the bookmark Word puts around the field. Word leaves it empty or repeats it at times. */
	name?: string;
	/** `w:enabled`; false when "Fill-in enabled" is off. Unset when the element is absent. */
	enabled?: boolean;
	/** `w:calcOnExit`: update the document's fields when the user leaves this one. */
	calcOnExit?: boolean;
	helpText?: FormFieldInfoText;
	statusText?: FormFieldInfoText;
	/** `w:textInput`: what a text field accepts and its default text. */
	textInput?: { type?: StFFTextType; default?: string; maxLength?: number; format?: string };
	/** `w:checkBox`: `checked` unset means the box shows its default state. Size is in half-points. */
	checkBox?: { default: boolean; checked?: boolean; sizeAuto?: boolean; size?: number };
	/** `w:ddList`: the entries, and the index of the selected (`result`) and default entries. */
	dropDown?: { entries: string[]; result?: number; default?: number };
	/** Verbatim `w:ffData`, used to keep what is not modeled when the field is written back. */
	sourceXml?: string;
}

/** The password verifier of `w:documentProtection`, in either spelling Word writes. */
export interface ProtectionPassword {
	/** The digest: `w:algorithmName`, or the one `w:cryptAlgorithmSid` names. */
	algorithmName: string;
	/** Base64 `w:hashValue` or `w:hash`. */
	hashValue: string;
	/** Base64 `w:saltValue` or `w:salt`. */
	saltValue: string;
	/** `w:spinCount` or `w:cryptSpinCount`. */
	spinCount?: number;
	/** Set for the Word form (`w:crypt*` attributes), which is also how it is written back. */
	word?: { algorithmSid: number };
}

/** `w:documentProtection` in settings.xml: what the document lets a user change. */
export interface DocumentProtection {
	/** `w:edit`: `forms` allows filling form fields and content controls only. */
	edit: StDocProtect;
	/** `w:enforcement`: whether the protection applies; Word enforces it unless it is off. */
	enforced: boolean;
	/** `w:formatting`: formatting is limited to the styles the document allows. */
	formatting?: boolean;
	/** The password that stops the protection. */
	password?: ProtectionPassword;
	/** True when the element has a password this editor cannot verify (an old-style hash). */
	unverifiablePassword?: boolean;
}

/** One editable range (`w:permStart` to `w:permEnd`) clipped to a paragraph, in characters. */
export interface PermissionRange {
	/** `w:id`, shared by the start and end markers. */
	id: string;
	/** `w:edGrp`: the group allowed to edit (`everyone` for Word's Exceptions list). */
	group?: string;
	/** `w:ed`: a single user allowed to edit. */
	user?: string;
	/** Character positions in the paragraph (a field, picture or break counts as one). */
	from: number;
	to: number;
	/** The range began in an earlier paragraph, or goes on into the next: the pane leaves it alone. */
	fromPrevious?: true;
	toNext?: true;
}
