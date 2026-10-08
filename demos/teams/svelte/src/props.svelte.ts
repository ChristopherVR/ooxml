// The props the demo mounts the binding with, as a `$state` object so changing a field re-renders
// it (a plain `.ts` module cannot use runes). The browser tests swap the host class
// (../../test-hooks).
import { config, userId, userName, workspaceId } from '../../shared';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

export const props = $state({
	class: currentHostClass(),
	workspaceId,
	userName,
	userId,
	config,
	onOpenFile: recordOpenFile,
});
onHostClass((value) => (props.class = value));
