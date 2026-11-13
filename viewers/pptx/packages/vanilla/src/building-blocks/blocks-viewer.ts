import type { HostDialogRequest } from '../viewer/editor/editor-controller';
import { PptxViewer } from '../viewer/PptxViewer';
import type { PptxViewerOptions } from '../viewer/types';

/**
 * The viewer behind the building blocks' canvas: the full vanilla engine with its chrome turned
 * off, whose dialogs (hyperlink, equation, find) become host requests instead of opening.
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

	override destroy(): void {
		this.onHostDialog = undefined;
		super.destroy();
	}
}
