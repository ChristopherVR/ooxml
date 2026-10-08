import { mergeMessageModules } from '../../../i18n/index';

/**
 * Merges the string files of one locale folder (`shell.ts`, `grid.ts`, `commands.ts`, ...). Each
 * file may export its table under any name (or as default); every exported plain object whose
 * values are all strings is merged, in file-name order, so agents can add files independently.
 */
export const mergeStringModules = mergeMessageModules;
