<script lang="ts">
	import { hasTextProperties } from 'pptx-viewer-core';
	import {
		fontHomeControls,
		homeSnapshotTranslator,
		tableTextShadowUpdate,
		textColorOf,
		transformInlineListCase,
	} from 'ooxml-ui/pptx';
	import type { ChangeCaseMode, RibbonHomeRequestEvent } from 'ooxml-ui/pptx';
	import { useTranslator } from '../../../../i18n/context';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import FontPickerGroup from './FontPickerGroup.svelte';
	import {
		tableTextChangeCasePatch,
		tableTextFormatOf,
		tableTextFormatPatch,
	} from '../../../editor/editor-table-text-format';
	import {
		adjustFontSizePatch,
		changeCasePatch,
		clearFormattingPatch,
		hasTextShadow,
		highlightColorOf,
		setCharacterSpacingPatch,
		setHighlightColorPatch,
		setTextColorPatch,
		toggleStrikethroughPatch,
		toggleTextFlagPatch,
		toggleTextShadowPatch,
	} from '../../../editor';
	
	/**
	 * The whole Home Font strip is the shared `pptx-ui-ribbon-home-font` element:
	 * character toggles, Text Shadow, size steps, Clear Formatting, Character
	 * Spacing, Change Case and the font and highlight colour popovers. Every
	 * edit still goes through `EditorState.patchSelected`.
	 */
	const { editor }: { editor: EditorState } = $props();
	const t = useTranslator();

	const element = $derived(editor.selectedElement);
	// A selected table formats every run of its selected cells (PowerPoint).
	const table = $derived(tableTextFormatOf(editor, element));
	const active = $derived(
		editor.editable && element !== undefined && (hasTextProperties(element) || table !== undefined),
	);
	const style = $derived(element && hasTextProperties(element) ? element.textStyle : undefined);
	const state = $derived({
		controls: fontHomeControls({
			enabled: active,
			bold: Boolean(table?.bold ?? style?.bold),
			italic: Boolean(table?.italic ?? style?.italic),
			underline: Boolean(table?.underline ?? style?.underline),
			strikethrough: Boolean(table?.strikethrough ?? style?.strikethrough),
			shadow: table ? Boolean(table.textShadowColor) : hasTextShadow(element),
			characterSpacing: table ? table.characterSpacing : style?.characterSpacing,
			fontColor: {
				value: table?.color ?? (element ? textColorOf(element) : '#000000'),
				ref: table ? undefined : style?.colorRef,
				themeColors: editor.themeColorMap,
				recent: editor.mruColors,
			},
			highlight: {
				value: (table?.highlightColor ?? (element ? highlightColorOf(element) : '')) || '#ffff00',
				recent: editor.mruColors,
			},
		}),
		translate: homeSnapshotTranslator(['font'], t),
	});

	/** Table cells take the character formats through the shared range op. */
	function requestTable(name: string, value: unknown): boolean {
		if (!table) {
			return false;
		}
		if (name === 'changeCase') {
			const mode = value as ChangeCaseMode;
			editor.patchSelected((current) => tableTextChangeCasePatch(editor, current, mode) ?? {});
			return true;
		}
		const size = (pt: number) => ({ fontSize: Math.min(400, Math.max(1, pt)) });
		const build = (current: NonNullable<typeof table>) => {
			switch (name) {
				case 'bold':
				case 'italic':
				case 'underline':
				case 'strikethrough':
					return { [name]: !current[name] };
				case 'shadow':
					return tableTextShadowUpdate(Boolean(current.textShadowColor));
				case 'clearFormatting':
					return {
						bold: false,
						italic: false,
						underline: false,
						strikethrough: false,
						highlightColor: undefined,
						characterSpacing: 0,
						...tableTextShadowUpdate(true),
					};
				case 'increaseFontSize':
					return size((current.fontSize ?? 18) + 2);
				case 'decreaseFontSize':
					return size((current.fontSize ?? 18) - 2);
				case 'fontColor':
					return { color: String(value) };
				case 'highlightColor':
					return { highlightColor: String(value) };
				case 'characterSpacing':
					return { characterSpacing: Number(value) };
				default:
					return undefined;
			}
		};
		if (build(table) === undefined) {
			return false;
		}
		editor.patchSelected((current) => tableTextFormatPatch(editor, current, (now) => build(now) ?? {}) ?? {});
		if (name === 'fontColor' || name === 'highlightColor') {
			editor.recordRecentColor(String(value));
		}
		return true;
	}

	function request(event: RibbonHomeRequestEvent): void {
		const { id, value, ref } = event.detail;
		if (requestTable(id.replace('home.font.', ''), value)) {
			return;
		}
		switch (id.replace('home.font.', '')) {
			case 'bold':
			case 'italic':
			case 'underline': {
				const flag = id.replace('home.font.', '') as 'bold' | 'italic' | 'underline';
				editor.patchSelected((current) => toggleTextFlagPatch(current, flag));
				break;
			}
			case 'strikethrough':
				editor.patchSelected((current) => toggleStrikethroughPatch(current));
				break;
			case 'shadow':
				editor.patchSelected((current) => toggleTextShadowPatch(current));
				break;
			case 'increaseFontSize':
				editor.patchSelected((current) => adjustFontSizePatch(current, 2));
				break;
			case 'decreaseFontSize':
				editor.patchSelected((current) => adjustFontSizePatch(current, -2));
				break;
			case 'clearFormatting':
				editor.patchSelected((current) => clearFormattingPatch(current));
				break;
			case 'characterSpacing':
				editor.patchSelected((current) => setCharacterSpacingPatch(current, Number(value)));
				break;
			case 'changeCase': {
				const mode = value as Parameters<typeof changeCasePatch>[1];
				editor.patchSelected((current, snapshot) => {
					if (!snapshot) {return changeCasePatch(current, mode);}
					const { text, textSegments } = transformInlineListCase(snapshot, null, mode);
					return { text, textSegments };
				});
				break;
			}
			case 'fontColor':
				editor.patchSelected((current) => setTextColorPatch(current, String(value), ref));
				editor.recordRecentColor(String(value));
				break;
			case 'highlightColor':
				editor.patchSelected((current) => setHighlightColorPatch(current, String(value)));
				editor.recordRecentColor(String(value));
		}
	}
</script>

<div data-pptx-chrome="font-controls">
	<FontPickerGroup {editor} />
	<pptx-ui-ribbon-home-font {state} onhome-request={request}></pptx-ui-ribbon-home-font>
</div>
