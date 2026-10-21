import type { EditorCore } from './editor-core';
import { applyLink, linkAtSelection, removeLink } from './link-commands';
import { SearchController, type SearchStatus } from './search-controller';

/** What a host dialog gets to edit the link at the selection. The host draws the dialog. */
export interface HyperlinkRequest {
	/** The link target at the selection, or null when the selection is not a link. */
	readonly href: string | null;
	/** The selected text, which becomes the link text. */
	readonly text: string;
	/** Applies `href`; `text` replaces the selection (or inserts at the caret) when given. */
	apply(href: string, text?: string): boolean;
	remove(): boolean;
}

/** Find and replace over the document. Find works while read-only; replace needs edit permission. */
export interface FindReplaceRequest {
	readonly status: SearchStatus;
	readonly canReplace: boolean;
	search(query: string, caseSensitive?: boolean): SearchStatus;
	findNext(): boolean;
	findPrevious(): boolean;
	replaceOne(replacement: string): boolean;
	replaceAll(replacement: string): number;
}

/** A host picture picker hands the chosen file back through `insert`. */
export interface ImagePickerRequest {
	insert(file: File): Promise<boolean>;
}

/**
 * Host-owned workflows for the surface. Each is named like the PowerPoint building blocks'
 * `onOpen...` callbacks. A workflow without its callback reports as not enabled; the library never
 * mounts a dialog of its own in surface mode.
 */
export interface SurfaceHostActions {
	onOpenHyperlinkDialog?: ((request: HyperlinkRequest) => void) | undefined;
	onOpenFindReplace?: ((request: FindReplaceRequest) => void) | undefined;
	onOpenImagePicker?: ((request: ImagePickerRequest) => void) | undefined;
}

function hyperlinkRequest(core: EditorCore): HyperlinkRequest | undefined {
	const view = core.targetView();
	if (!view) return undefined;
	const { from, to } = view.state.selection;
	return {
		href: linkAtSelection(view)?.href ?? null,
		text: view.state.doc.textBetween(from, to, ' '),
		apply: (href, text) => core.canEditBody() && applyLink(view, { href }, text),
		remove: () => core.canEditBody() && removeLink(view),
	};
}

function findReplaceRequest(core: EditorCore): FindReplaceRequest {
	const controller = new SearchController(() => core.targetView());
	return {
		get status() {
			return controller.status;
		},
		get canReplace() {
			return core.canEditBody();
		},
		search: (query, caseSensitive) => controller.search(query, caseSensitive),
		findNext: () => controller.findNext() !== null,
		findPrevious: () => controller.findPrevious() !== null,
		replaceOne: (replacement) => core.canEditBody() && controller.replaceOne(replacement),
		replaceAll: (replacement) => (core.canEditBody() ? controller.replaceAll(replacement) : 0),
	};
}

/** Enabled flags and openers for the host's link, find/replace and picture workflows. */
export function wordHostActionProps(core: EditorCore) {
	// Evaluated at call time, so a retained callback follows permission and host changes.
	const ready = () => Boolean(core.targetView()) && !core.loading;
	const editable = () => ready() && core.canEditBody() && Boolean(core.targetView()?.editable);
	const linkEnabled = () => editable() && Boolean(core.surfaceActions.onOpenHyperlinkDialog);
	const findEnabled = () => ready() && Boolean(core.surfaceActions.onOpenFindReplace);
	const pictureEnabled = () => editable() && Boolean(core.surfaceActions.onOpenImagePicker);
	const supplied = core.surfaceActions;
	return {
		/** Which workflows the host supplied a callback for, whether or not they are usable now. */
		hostWorkflows: {
			hyperlink: Boolean(supplied.onOpenHyperlinkDialog),
			findReplace: Boolean(supplied.onOpenFindReplace),
			picture: Boolean(supplied.onOpenImagePicker),
		},
		canInsertLink: linkEnabled(),
		canFindReplace: findEnabled(),
		canInsertPicture: pictureEnabled(),
		/** Opens the host's link dialog for the selection. Returns false when not enabled. */
		onOpenHyperlinkDialog: (): boolean => {
			const open = core.surfaceActions.onOpenHyperlinkDialog;
			const request = linkEnabled() && open ? hyperlinkRequest(core) : undefined;
			if (!request || !open) return false;
			open(request);
			return true;
		},
		/** Opens the host's find and replace workflow. Returns false when not enabled. */
		onOpenFindReplace: (): boolean => {
			const open = core.surfaceActions.onOpenFindReplace;
			if (!findEnabled() || !open) return false;
			open(findReplaceRequest(core));
			return true;
		},
		/** Opens the host's picture picker. Returns false when not enabled. */
		onOpenImagePicker: (): boolean => {
			const open = core.surfaceActions.onOpenImagePicker;
			if (!pictureEnabled() || !open) return false;
			open({
				insert: (file) =>
					editable() ? core.inserts.insertPictureFile(file) : Promise.resolve(false),
			});
			return true;
		},
	};
}
