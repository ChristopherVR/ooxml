import { combo, command, group, menu, stack, type CommandSpec } from './ribbon-parts.js';

const TEXT = 'Needs core text formatting edits.';
const CLIPBOARD = 'Needs core shape copy and paste.';
const STYLE = 'Needs core fill, line and effect edits.';
const small = (spec: CommandSpec): CommandSpec => ({ size: 'small', ...spec });
const icon = (spec: CommandSpec): CommandSpec => ({ size: 'icon', ...spec });

/** Visio's Home > Clipboard, Font and Paragraph groups, in Visio's layout. */
export function textGroups(doc: Document): [HTMLElement, HTMLElement, HTMLElement] {
	const clipboard = group(doc, 'Clipboard', [
		menu(doc, {
			id: 'paste',
			label: 'Paste',
			icon: 'paste',
			split: true,
			unsupported: CLIPBOARD,
			keys: ['Control+V', 'Ctrl+V'],
			items: [
				{ id: 'paste-item', label: 'Paste', unsupported: CLIPBOARD },
				{ id: 'paste-special', label: 'Paste Special...', unsupported: CLIPBOARD },
			],
		}),
		stack(doc, [
			command(doc, small({ id: 'cut', label: 'Cut', icon: 'cut', unsupported: CLIPBOARD })),
			command(doc, small({ id: 'copy', label: 'Copy', icon: 'copy', unsupported: CLIPBOARD })),
			command(
				doc,
				small({
					id: 'format-painter',
					label: 'Format Painter',
					icon: 'formatPainter',
					unsupported: STYLE,
				}),
			),
		]),
	]);
	const font = group(
		doc,
		'Font',
		[
			stack(doc, [
				stack(
					doc,
					[
						combo(doc, {
							id: 'font',
							label: 'Font',
							placeholder: 'Calibri',
							width: 128,
							unsupported: TEXT,
						}),
						combo(doc, {
							id: 'font-size',
							label: 'Font Size',
							placeholder: '12pt.',
							width: 64,
							unsupported: TEXT,
						}),
						command(
							doc,
							icon({
								id: 'grow-font',
								label: 'Increase Font Size',
								icon: 'growFont',
								unsupported: TEXT,
							}),
						),
						command(
							doc,
							icon({
								id: 'shrink-font',
								label: 'Decrease Font Size',
								icon: 'shrinkFont',
								unsupported: TEXT,
							}),
						),
					],
					true,
				),
				stack(
					doc,
					[
						command(doc, icon({ id: 'bold', label: 'Bold', icon: 'bold', unsupported: TEXT })),
						command(
							doc,
							icon({ id: 'italic', label: 'Italic', icon: 'italic', unsupported: TEXT }),
						),
						command(
							doc,
							icon({ id: 'underline', label: 'Underline', icon: 'underline', unsupported: TEXT }),
						),
						command(
							doc,
							icon({
								id: 'strikethrough',
								label: 'Strikethrough',
								icon: 'strikethrough',
								unsupported: TEXT,
							}),
						),
						menu(doc, {
							id: 'change-case',
							label: 'Change Case',
							icon: 'changeCase',
							size: 'icon',
							unsupported: TEXT,
							items: [{ id: 'upper', label: 'UPPERCASE', unsupported: TEXT }],
						}),
						menu(doc, {
							id: 'font-color',
							label: 'Font Color',
							icon: 'fontColor',
							size: 'icon',
							split: true,
							unsupported: TEXT,
							items: [{ id: 'more-colors', label: 'More Colors...', unsupported: TEXT }],
						}),
					],
					true,
				),
			]),
		],
		{ launcher: TEXT },
	);
	const paragraph = group(
		doc,
		'Paragraph',
		[
			stack(doc, [
				stack(
					doc,
					[
						command(
							doc,
							icon({ id: 'align-top', label: 'Align Top', icon: 'alignTop', unsupported: TEXT }),
						),
						command(
							doc,
							icon({
								id: 'align-middle',
								label: 'Align Middle',
								icon: 'alignMiddle',
								unsupported: TEXT,
							}),
						),
						command(
							doc,
							icon({
								id: 'align-bottom',
								label: 'Align Bottom',
								icon: 'alignBottom',
								unsupported: TEXT,
							}),
						),
						command(
							doc,
							icon({ id: 'bullets', label: 'Bullets', icon: 'bullets', unsupported: TEXT }),
						),
					],
					true,
				),
				stack(
					doc,
					[
						command(
							doc,
							icon({ id: 'align-left', label: 'Align Left', icon: 'alignLeft', unsupported: TEXT }),
						),
						command(
							doc,
							icon({ id: 'align-center', label: 'Center', icon: 'alignCenter', unsupported: TEXT }),
						),
						command(
							doc,
							icon({
								id: 'align-right',
								label: 'Align Right',
								icon: 'alignRight',
								unsupported: TEXT,
							}),
						),
						command(
							doc,
							icon({ id: 'justify', label: 'Justify', icon: 'justify', unsupported: TEXT }),
						),
						command(
							doc,
							icon({
								id: 'indent-decrease',
								label: 'Decrease Indent',
								icon: 'indentDecrease',
								unsupported: TEXT,
							}),
						),
						command(
							doc,
							icon({
								id: 'indent-increase',
								label: 'Increase Indent',
								icon: 'indentIncrease',
								unsupported: TEXT,
							}),
						),
					],
					true,
				),
			]),
		],
		{ launcher: TEXT },
	);
	return [clipboard, font, paragraph];
}
