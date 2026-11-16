// The strings of the shared SmartArt UI (Insert SmartArt dialog, SmartArt Design tab, text pane,
// the insert catalogue's category and layout names, and the Change Colors and SmartArt Styles
// gallery names) in every product locale, moved from Word's `src/ui/src/docx/locales/smartart.ts`
// so Word and Excel show the same translations. Keys are the English UI text; catalogue names use
// `smartart.*` keys so they never collide with other text.
import { de } from './de';
import { en as enUi, type DiagramUiStrings as UiStrings } from './en';
import { es } from './es';
import { fr } from './fr';
import { galleryEn, type GalleryStrings } from './gallery';
import { galleryDe, galleryEs, galleryFr, galleryZhCN } from './gallery-locales';
import { zhCN } from './zh-CN';

/** Every SmartArt UI string key; each locale has them all. */
export type DiagramUiStrings = UiStrings & GalleryStrings;

const en = { ...enUi, ...galleryEn } as const;

/** The SmartArt UI strings by locale; English keeps its literal values (keys of dictionaries). */
export const DIAGRAM_UI_STRINGS: {
	readonly en: typeof en;
	readonly fr: DiagramUiStrings;
	readonly de: DiagramUiStrings;
	readonly es: DiagramUiStrings;
	readonly 'zh-CN': DiagramUiStrings;
} = {
	en,
	fr: { ...fr, ...galleryFr },
	de: { ...de, ...galleryDe },
	es: { ...es, ...galleryEs },
	'zh-CN': { ...zhCN, ...galleryZhCN },
};
