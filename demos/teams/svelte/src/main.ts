import { mount } from 'svelte';
import Teams from 'openteams-svelte-viewer';
import { showStaticNotice } from '../../shared';
import { props } from './props.svelte';

showStaticNotice();
mount(Teams, { target: document.getElementById('app')!, props });
