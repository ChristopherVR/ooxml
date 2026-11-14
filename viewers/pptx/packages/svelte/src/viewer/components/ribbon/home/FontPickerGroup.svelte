<script lang="ts">
	/**
	 * FontPickerGroup: the Home > Font family and size selects as the shared
	 * `pptx-ui-ribbon-home-font-picker` element. The element renders the group
	 * wrapper, caption and both `pptx-ui-select` fields; this adapter supplies
	 * the deck's font lists and the current values and applies picked values
	 * through `EditorState.patchSelected`.
	 */
	import { hasTextProperties } from 'pptx-viewer-core';
	import {
		fontPickerHomeControls,
		fontSizeOf,
		homeSnapshotTranslator,
		resolveDefaultFontFamily,
	} from 'ooxml-ui/pptx';
	import type { RibbonHomeRequestEvent } from 'ooxml-ui/pptx';

	import { useTranslator } from '../../../../i18n/context';
	import { setFontFamilyPatch, setFontSizePatch } from '../../../editor';
	import { tableTextFormatOf, tableTextFormatPatch } from '../../../editor/editor-table-text-format';
	import type { EditorState } from '../../../editor/editor-state.svelte';

	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();
	const el = $derived(editor.selectedElement);
	// A selected table sets the font of every run of its selected cells.
	const table = $derived(tableTextFormatOf(editor, el));
	const active = $derived(el !== undefined && (hasTextProperties(el) || table !== undefined));
	const themeFonts = $derived({
		heading: editor.theme?.fontScheme?.majorFont?.latin,
		body: editor.theme?.fontScheme?.minorFont?.latin,
	});
	// With nothing overriding it, the box shows the family the deck would render.
	const fontFamily = $derived(
		table?.fontFamily ??
		(el && hasTextProperties(el) ? el.textStyle?.fontFamily : undefined) ??
			resolveDefaultFontFamily(
				(el as { placeholderType?: string } | undefined)?.placeholderType,
				themeFonts,
			),
	);
	const state = $derived({
		controls: fontPickerHomeControls(
			{
				enabled: editor.editable && active,
				fontFamily,
				fontSize: table?.fontSize ?? (el ? fontSizeOf(el) : 24),
				themeFonts,
				embeddedFonts: editor.embeddedFontFamilies,
				customFonts: editor.customFontFamilies,
			},
			t,
		),
		translate: homeSnapshotTranslator(['font-picker'], t),
	});

	function request(event: RibbonHomeRequestEvent): void {
		const { id, value } = event.detail;
		if (!editor.editable || !active) {
			return;
		}
		if (table) {
			const size = Number(value);
			const update =
				id === 'home.font.fontFamily'
					? { fontFamily: String(value) }
					: id === 'home.font.fontSize' && Number.isFinite(size) && size > 0
						? { fontSize: size }
						: undefined;
			if (update) {
				editor.patchSelected((current) => tableTextFormatPatch(editor, current, () => update) ?? {});
			}
			return;
		}
		if (id === 'home.font.fontFamily') {
			editor.patchSelected((current) => setFontFamilyPatch(current, String(value)));
		} else if (id === 'home.font.fontSize') {
			const next = Number(value);
			if (Number.isFinite(next) && next > 0) {
				editor.patchSelected((current) => setFontSizePatch(current, next));
			}
		}
	}
</script>

<pptx-ui-ribbon-home-font-picker {state} onhome-request={request}></pptx-ui-ribbon-home-font-picker>
