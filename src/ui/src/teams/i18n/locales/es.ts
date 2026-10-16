import { esCalendar } from './es/calendar';
import { esChat } from './es/chat';
import { esCore } from './es/core';
import { esFiles } from './es/files';
import { esSettings } from './es/settings';
import { esShell } from './es/shell';

/** Español. */
export const es = { ...esShell, ...esChat, ...esCalendar, ...esFiles, ...esSettings, ...esCore };
