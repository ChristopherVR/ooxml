import { createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { Teams } from 'openteams-solid-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

// The host class is a reactive prop; the browser tests swap it (../../test-hooks).
const [hostClass, setHostClass] = createSignal(currentHostClass());
onHostClass(setHostClass);

showStaticNotice();
render(
	() =>
		Teams({
			get class() {
				return hostClass();
			},
			workspaceId,
			userName,
			userId,
			config,
			onOpenFile: recordOpenFile,
		}),
	document.getElementById('app')!,
);
