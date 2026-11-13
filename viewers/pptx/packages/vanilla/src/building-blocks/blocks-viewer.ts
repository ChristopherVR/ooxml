import type { BuildingBlockLiveOptionPort, GoogleFontsOption } from 'ooxml-ui/pptx';

import type { HostDialogRequest } from '../viewer/editor/editor-controller';
import { PptxViewer } from '../viewer/PptxViewer';
import type { PptxViewerOptions } from '../viewer/types';

/**
 * The viewer behind the building blocks' canvas: the full vanilla engine with its chrome turned
 * off, whose dialogs (hyperlink, equation, find) become host requests instead of opening. It also
 * takes the options the vanilla viewer otherwise reads once (`googleFonts`, `allowExternalContent`,
 * `collaboration`) as live changes, through {@link liveOptionPort}.
 */
export class BuildingBlocksViewer extends PptxViewer {
	private onHostDialog: ((request: HostDialogRequest) => boolean) | undefined;

	constructor(
		container: HTMLElement,
		options: PptxViewerOptions,
		onHostDialog: (request: HostDialogRequest) => boolean,
	) {
		super(container, options);
		this.onHostDialog = onHostDialog;
	}

	protected override requestHostDialog(request: HostDialogRequest): boolean {
		return this.onHostDialog?.(request) ?? false;
	}

	/** The collaboration session's role or readiness blocks editing. */
	collaborationReadOnly(): boolean {
		return this.isCollaborationReadOnly();
	}

	/** Apply a live option change to the mounted viewer (the open deck stays). */
	readonly liveOptionPort: BuildingBlockLiveOptionPort = {
		startCollaboration: (input) => this.startCollaboration(input),
		stopCollaboration: () => this.stopCollaboration(),
		setGoogleFonts: (option: GoogleFontsOption | undefined) => {
			// Locks Trust Center > "Download fonts"; the webfonts follow the new policy at once.
			this.customization.controller.setGoogleFonts(option);
			this.refreshWebfonts();
		},
		setAllowExternalContent: (allow) => {
			this.customization.controller.setAllowExternalContent(allow);
		},
	};

	override destroy(): void {
		this.onHostDialog = undefined;
		super.destroy();
	}
}
