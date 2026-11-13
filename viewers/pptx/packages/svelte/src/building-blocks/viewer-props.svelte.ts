import type { PowerPointViewerProps } from '../viewer/types';
import type { ViewerBuildingBlocksOptions } from './types';

/**
 * The viewer props the blocks hand to the viewer state (the blocks mount no chrome), as getters
 * over the live options. `content` and `canEdit` are `$derived`, so the load effect only re-runs
 * when the deck itself changes, not on every option update.
 */
export function viewerProps(getOptions: () => ViewerBuildingBlocksOptions): PowerPointViewerProps {
	const content = $derived(getOptions().content);
	const canEdit = $derived(getOptions().canEdit === true);
	const option = <K extends keyof ViewerBuildingBlocksOptions>(key: K) => getOptions()[key];
	return {
		get source() {
			return content;
		},
		get editable() {
			return canEdit;
		},
		// Recovery snapshots are the host's to opt into; the blocks never write them by default.
		get autosave() {
			return option('autosave') === true;
		},
		get fileName() {
			return option('fileName');
		},
		get filePath() {
			return option('filePath');
		},
		get initialSlide() {
			return option('initialSlide');
		},
		get googleFonts() {
			return option('googleFonts');
		},
		get listConversionScope() {
			return option('listConversionScope');
		},
		get allowExternalContent() {
			return option('allowExternalContent');
		},
		get limits() {
			return option('limits');
		},
		get collaboration() {
			return option('collaboration');
		},
		get fitPadding() {
			return option('fitPadding');
		},
		get maxFitScale() {
			return option('maxFitScale');
		},
		get authorName() {
			return option('authorName') ?? option('userName');
		},
		get onload() {
			return option('onLoad');
		},
		get onerror() {
			return option('onError');
		},
		get onslidechange() {
			return option('onSlideChange');
		},
		get ondirtychange() {
			return option('onDirtyChange');
		},
		get oncontentchange() {
			return option('onContentChange');
		},
		get onselectionchange() {
			return option('onSelectionChange');
		},
	} as PowerPointViewerProps;
}
