/**
 * The host `allowExternalContent` switch: whether a deck may make the browser
 * contact other servers for linked (`TargetMode="External"`, http(s)) images
 * and media.
 *
 * Unset keeps the user's File > Options > Trust Center choice
 * (`trust.allowExternalContent`). A boolean overrides it for every load and
 * locks the Trust Center toggle to that value, so the user cannot change it.
 * Every binding feeds the result of {@link resolveAllowExternalContent} to the
 * loader's `allowExternalImages` option and folds
 * {@link lockExternalContentSetting} into the options locks; neither rule is
 * re-derived per binding.
 *
 * @module render/options/external-content-policy
 */
import type { OptionsSettingId, OptionsSettingValues } from '../customization/customization-types';
import type { ViewerOptions } from './viewer-options';

/** The `<group>.<key>` id of the Trust Center setting the host switch locks. */
export const EXTERNAL_CONTENT_SETTING_ID: OptionsSettingId = 'trust.allowExternalContent';

/**
 * Whether a load may fetch linked external content: the host value when set,
 * otherwise the user's Trust Center option.
 */
export function resolveAllowExternalContent(
	host: boolean | undefined,
	options: Pick<ViewerOptions, 'trust'>,
): boolean {
	return typeof host === 'boolean' ? host : options.trust.allowExternalContent;
}

/**
 * The options locks with the host switch folded in. Returns `locked` itself
 * when the host left the switch unset, and a new record otherwise; the host
 * switch wins over a `customization.options.locked` entry for the same id.
 */
export function lockExternalContentSetting(
	locked: OptionsSettingValues,
	host: boolean | undefined,
): OptionsSettingValues {
	if (typeof host !== 'boolean') {
		return locked;
	}
	return { ...locked, [EXTERNAL_CONTENT_SETTING_ID]: host };
}
