import { frCalendar } from './fr/calendar';
import { frChat } from './fr/chat';
import { frCore } from './fr/core';
import { frFiles } from './fr/files';
import { frSettings } from './fr/settings';
import { frShell } from './fr/shell';

/** Français. */
export const fr = { ...frShell, ...frChat, ...frCalendar, ...frFiles, ...frSettings, ...frCore };
