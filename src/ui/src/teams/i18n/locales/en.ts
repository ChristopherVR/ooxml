import { enCalendar } from './en/calendar';
import { enChat } from './en/chat';
import { enCore } from './en/core';
import { enFiles } from './en/files';
import { enSettings } from './en/settings';
import { enShell } from './en/shell';

/** The English dictionary: the contract every other locale must meet key for key. */
export const en = { ...enShell, ...enChat, ...enCalendar, ...enFiles, ...enSettings, ...enCore };
