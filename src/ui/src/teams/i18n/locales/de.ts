import { deCalendar } from './de/calendar';
import { deChat } from './de/chat';
import { deCore } from './de/core';
import { deFiles } from './de/files';
import { deSettings } from './de/settings';
import { deShell } from './de/shell';

/** Deutsch. */
export const de = { ...deShell, ...deChat, ...deCalendar, ...deFiles, ...deSettings, ...deCore };
