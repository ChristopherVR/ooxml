// The strings of the shared SmartArt UI (Insert SmartArt dialog, SmartArt Design tab, text pane and
// the insert catalogue's category and layout names) in every product locale, moved from Word's
// `src/ui/src/docx/locales/smartart.ts` so Word and Excel show the same translations. Keys are the
// English UI text; catalogue names use `smartart.*` keys so they never collide with other text.
import { de } from './de';
import { en, type DiagramUiStrings } from './en';
import { es } from './es';
import { fr } from './fr';
import { zhCN } from './zh-CN';

export type { DiagramUiStrings };

/** The SmartArt UI strings by locale; English keeps its literal values (keys of dictionaries). */
export const DIAGRAM_UI_STRINGS: {
	readonly en: typeof en;
	readonly fr: DiagramUiStrings;
	readonly de: DiagramUiStrings;
	readonly es: DiagramUiStrings;
	readonly 'zh-CN': DiagramUiStrings;
} = { en, fr, de, es, 'zh-CN': zhCN };
