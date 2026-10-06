import { render } from 'solid-js/web';
import { Teams } from 'openteams-solid-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';

showStaticNotice();
render(() => Teams({ workspaceId, userName, userId, config }), document.getElementById('app')!);
