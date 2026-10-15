<script setup lang="ts">
import {
	DESIGN_RIBBON_COMMANDS,
	DESIGN_RIBBON_GROUPS,
	FIXED_TAB_GALLERIES,
	designCommandState,
} from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import RibbonGallery from './RibbonGallery.vue';

interface Props {
	canEdit: boolean;
	onToggleThemeGallery: () => void;
	isThemeGalleryOpen: boolean;
	onToggleThemeEditor: () => void;
	isThemeEditorOpen: boolean;
	onOpenDocumentProperties?: () => void;
	onOpenSlideSize?: () => void;
	onToggleInspector?: () => void;
	isInspectorPaneOpen?: boolean;
}
const props = defineProps<Props>();
const { t } = useI18n();
const actions = computed(() => ({
	'design.themes.browseThemes': props.onToggleThemeGallery,
	'design.themes.editTheme': props.onToggleThemeEditor,
	'design.customize.slideSize': props.onOpenSlideSize ?? props.onOpenDocumentProperties,
	'design.customize.formatBackground': props.onToggleInspector,
}));
const groups = computed(() =>
	DESIGN_RIBBON_GROUPS.map((group) => ({
		...group,
		galleries: FIXED_TAB_GALLERIES.filter((placement) =>
			placement.control.startsWith(`${group.id}.`),
		),
		commands: DESIGN_RIBBON_COMMANDS.filter((command) => command.id.startsWith(`${group.id}.`))
			.map((command) => ({
				...command,
				...designCommandState(command.id, {
					editable: props.canEdit,
					galleryOpen: props.isThemeGalleryOpen,
					editorOpen: props.isThemeEditorOpen,
					backgroundOpen: props.isInspectorPaneOpen,
					hasSlideSize: Boolean(actions.value['design.customize.slideSize']),
					hasBackground: Boolean(props.onToggleInspector),
				}),
			}))
			.filter((command) => !command.hidden),
	})),
);
</script>
<template>
	<pptx-ui-ribbon-group
		v-for="group in groups"
		:key="group.id"
		:label="t(group.labelKey)"
		:data-ribbon-group="group.id"
	>
		<RibbonGallery
			v-for="placement in group.galleries"
			:key="placement.control"
			:gallery="placement.gallery"
			:control="placement.control"
			:mode="placement.mode"
		/>
		<pptx-ui-ribbon-command
			v-for="command in group.commands"
			:key="command.id"
			:data-ribbon-control="command.id"
			:label="t(command.labelKey)"
			:title="t(command.titleKey)"
			:icon="command.icon"
			:disabled="command.disabled || undefined"
			:active="command.active || undefined"
			:expanded="command.expanded === undefined ? undefined : String(command.expanded)"
			@command-request="actions[command.id as keyof typeof actions]?.()"
		></pptx-ui-ribbon-command>
	</pptx-ui-ribbon-group>
</template>
