// The props the demo mounts the binding with, as a `$state` object so that changing a field
// re-renders (a plain `.ts` module cannot use runes). The e2e specs swap the host class through
// ../../test-hooks.
import { config, locale, userId, userName, workspaceId } from '../../shared';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

export const props = $state({
	class: currentHostClass(),
	workspaceId,
	userName,
	userId,
	locale,
	config,
	onOpenFile: recordOpenFile,
});
onHostClass((value) => (props.class = value));
