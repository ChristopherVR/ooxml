import type { EditorHandle } from '../../../viewers/docx/packages/bindings/src/index';
import type {
	BuildingBlockOptions,
	WordNavigationProps,
	WordStatusProps,
	WordToolbarProps,
} from '../../../viewers/docx/packages/bindings/src/building-blocks';
export interface NativeHarness {
	events?: { document: number; dirty: number; error: number };
	handle: EditorHandle;
	toolbar: () => WordToolbarProps | null;
	navigation: () => WordNavigationProps | null;
	status: () => WordStatusProps | null;
	permission: (canEdit: boolean) => void;
	options: (options: BuildingBlockOptions) => void;
	destroy: () => void | Promise<void>;
	error: () => Error | null;
	mode: () => string;
}
declare global {
	interface Window {
		nativeWord?: NativeHarness;
		remountNativeWord?: () => Promise<void>;
	}
}
