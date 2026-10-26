import type { PptxHandler } from 'pptx-viewer-core';

export const LOAD_LIMIT_DECK_SLIDES: 3;
export const OVER_SLIDE_LIMIT_MESSAGE: string;
export function loadLimitDeck(Handler: typeof PptxHandler): Promise<Uint8Array>;
