import type { HyperlinkEditData } from 'ooxml-ui/pptx';

// The editor values are shared by every binding's building blocks (`ooxml-ui/pptx`).
export type { HyperlinkActionVerb, HyperlinkEditData, HyperlinkTargetType } from 'ooxml-ui/pptx';
export {
	HYPERLINK_ACTION_VERB_MAP as ACTION_VERB_MAP,
	HYPERLINK_ACTION_VERB_TO_PPACTION as ACTION_VERB_TO_PPACTION,
} from 'ooxml-ui/pptx';

export interface HyperlinkEditDialogProps {
	open: boolean;
	initialUrl?: string;
	initialTooltip?: string;
	initialAction?: string;
	slideCount: number;
	onConfirm: (data: HyperlinkEditData) => void;
	onCancel: () => void;
}
