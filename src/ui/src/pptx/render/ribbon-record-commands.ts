import type { RibbonGroupView } from './ribbon-command-view';
import type { PptxSlide } from 'ooxml-core/pptx';
import { containsCameraClip } from 'ooxml-core/pptx/editor/render/slide-operations';
import { slideHasCameo, type RibbonGraphicTarget } from './ribbon-insert-state';
import { ribbonGroupViews, type SlideShowRibbonGroup } from './ribbon-slide-show-commands';

/** Shared presentation metadata; recording operations remain in the bindings. */
export const RECORD_COMMAND_GROUPS: readonly SlideShowRibbonGroup[] = [
	{
		id: 'record.camera',
		labelKey: 'pptx.record.camera',
		commands: [
			{
				id: 'record.camera.cameo',
				labelKey: 'pptx.record.cameo',
				icon: 'camera',
				unsupported: true,
			},
		],
	},
	{
		id: 'record.record',
		labelKey: 'pptx.ribbon.tab.record',
		commands: [
			{
				id: 'record.record.fromBeginning',
				labelKey: 'pptx.slideShow.fromBeginning',
				icon: 'video',
			},
			{ id: 'record.record.fromCurrent', labelKey: 'pptx.slideShow.fromCurrent', icon: 'play' },
		],
	},
	{
		id: 'record.manage',
		labelKey: 'pptx.record.manage',
		commands: [
			{
				id: 'record.manage.clear',
				labelKey: 'pptx.record.clearTimings',
				icon: 'eraser',
				unsupported: true,
			},
			{
				id: 'record.manage.reset',
				labelKey: 'pptx.record.resetToCameo',
				icon: 'reset',
				unsupported: true,
			},
		],
	},
	{
		id: 'record.help',
		labelKey: 'pptx.ribbon.tab.help',
		commands: [
			{
				id: 'record.help.learnMore',
				labelKey: 'pptx.record.learnMore',
				icon: 'help',
				unsupported: true,
			},
		],
	},
];

/** The Record tab for `pptx-ui-ribbon-section`; the bindings keep the recording itself. */
export type RecordScope = 'current' | 'all';
export type RecordScopeCommand =
	| 'record.camera.cameo'
	| 'record.manage.clear'
	| 'record.manage.reset';

export type RecordTarget = Pick<
	RibbonGraphicTarget,
	'slides' | 'slideIndex' | 'canvasSize' | 'insert' | 'commitSlides'
>;

export function recordScopeAvailable(
	id: RecordScopeCommand,
	target: RecordTarget,
	scope: RecordScope,
): boolean {
	const canApply = (slide: PptxSlide) => {
		if (id === 'record.camera.cameo') return !slideHasCameo(slide);
		if (id === 'record.manage.reset') return containsCameraClip(slide.elements);
		return slide.transition?.advanceAfterMs !== undefined;
	};
	if (id !== 'record.camera.cameo' && !target.commitSlides) return false;
	const current = target.slides[target.slideIndex];
	return scope === 'all'
		? !!target.commitSlides && target.slides.some(canApply)
		: !!current && canApply(current);
}

export function buildRecordRibbon(
	t: (key: string) => string,
	target?: RecordTarget,
): RibbonGroupView[] {
	return ribbonGroupViews(t, RECORD_COMMAND_GROUPS).map((group) => ({
		...group,
		commands: group.commands.map((command) => {
			const id = command.id;
			if (
				id !== 'record.camera.cameo' &&
				id !== 'record.manage.clear' &&
				id !== 'record.manage.reset'
			)
				return command;
			return {
				...command,
				disabled:
					!target ||
					(id === 'record.camera.cameo'
						? !target.slides[target.slideIndex]
						: !recordScopeAvailable(id, target, 'current') &&
							!recordScopeAvailable(id, target, 'all')),
			};
		}),
	}));
}
