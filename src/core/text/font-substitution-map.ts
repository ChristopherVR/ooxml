/** Ordered Office font substitutes. Extracted from the PowerPoint renderer. */
export const FONT_SUBSTITUTION_MAP: Record<string, readonly string[]> = {
	// Microsoft Office default fonts. The Aptos family became the Office
	// default and its metrics were drawn to sit close to Calibri, so falling
	// back through Calibri keeps authored line counts; the generic sans-serif
	// chain is wider and re-wraps one-line titles.
	Aptos: ['Calibri', 'Carlito', 'Liberation Sans', 'Arial', 'sans-serif'],
	'Aptos Display': ['Calibri Light', 'Calibri', 'Carlito', 'Arial', 'sans-serif'],
	'Aptos Narrow': ['Calibri', 'Liberation Sans Narrow', 'Arial', 'sans-serif'],
	'Aptos Serif': ['Cambria', 'Caladea', 'Liberation Serif', 'Times New Roman', 'serif'],
	'Aptos Mono': ['Consolas', 'Liberation Mono', 'Courier New', 'monospace'],
	Calibri: ['Carlito', 'Liberation Sans', 'Arial', 'sans-serif'],
	'Calibri Light': ['Carlito', 'Liberation Sans', 'Arial', 'sans-serif'],
	Cambria: ['Caladea', 'Liberation Serif', 'Times New Roman', 'serif'],
	'Cambria Math': ['STIX Two Math', 'Latin Modern Math', 'Times New Roman', 'serif'],
	Consolas: ['Liberation Mono', 'Courier New', 'monospace'],
	// Selawik is Microsoft's own metric-compatible open substitute for Segoe
	// UI (github.com/microsoft/Selawik); it is not on Google Fonts, so the
	// webfont loader (google-webfonts.ts) cannot fetch it automatically, but
	// listing it first still wins for a reader who has it installed or a
	// deployment that self-hosts it.
	'Segoe UI': ['Selawik', 'Liberation Sans', 'Arimo', 'Helvetica Neue', 'Arial', 'sans-serif'],
	'Segoe UI Light': [
		'Selawik Light',
		'Liberation Sans',
		'Arimo',
		'Helvetica Neue',
		'Arial',
		'sans-serif',
	],
	'Segoe UI Semibold': [
		'Selawik Semibold',
		'Liberation Sans',
		'Arimo',
		'Helvetica Neue',
		'Arial',
		'sans-serif',
	],

	// Classic Windows fonts → cross-platform equivalents. Arimo/Tinos/Cousine
	// are Google's metric-compatible clones of Arial/Times New Roman/Courier
	// New (the same project that produced Carlito/Caladea) and, unlike
	// Liberation Sans/Serif/Mono, are served by Google Fonts, so listing them
	// is what lets the webfont loader actually fetch a matching face when the
	// reader's system has neither the original nor the Liberation family.
	'Times New Roman': ['Liberation Serif', 'Tinos', 'Times', 'serif'],
	Arial: ['Liberation Sans', 'Arimo', 'Helvetica', 'sans-serif'],
	'Arial Black': ['Liberation Sans', 'Arimo', 'Helvetica', 'sans-serif'],
	'Arial Narrow': ['Liberation Sans Narrow', 'Arimo', 'Helvetica Neue', 'sans-serif'],
	Helvetica: ['Liberation Sans', 'Arimo', 'Arial', 'sans-serif'],
	'Helvetica Neue': ['Liberation Sans', 'Arimo', 'Arial', 'sans-serif'],
	'Courier New': ['Liberation Mono', 'Cousine', 'Courier', 'monospace'],
	Verdana: ['DejaVu Sans', 'Bitstream Vera Sans', 'sans-serif'],
	// Gelasio is Google's metric-compatible clone of Georgia (built the same
	// way as Carlito/Caladea from Georgia's own hmtx table); Liberation Serif
	// is not metrically close to Georgia and stays only as a further fallback.
	Georgia: ['Gelasio', 'Liberation Serif', 'Times New Roman', 'serif'],
	Tahoma: ['DejaVu Sans', 'Liberation Sans', 'sans-serif'],
	Trebuchet: ['Liberation Sans', 'sans-serif'],
	'Trebuchet MS': ['Liberation Sans', 'sans-serif'],
	'Comic Sans MS': ['Comic Neue', 'cursive'],
	Impact: ['Charcoal', 'sans-serif'],
	'Lucida Console': ['Liberation Mono', 'monospace'],
	'Lucida Sans Unicode': ['Lucida Grande', 'Liberation Sans', 'sans-serif'],
	Palatino: ['Palatino Linotype', 'Book Antiqua', 'serif'],
	'Palatino Linotype': ['Palatino', 'Book Antiqua', 'serif'],
	'Book Antiqua': ['Palatino Linotype', 'Palatino', 'serif'],

	// CJK fonts
	'MS PGothic': ['Noto Sans CJK JP', 'Hiragino Sans', 'Yu Gothic', 'sans-serif'],
	'MS PMincho': ['Noto Serif CJK JP', 'Hiragino Mincho ProN', 'Yu Mincho', 'serif'],
	'MS Gothic': ['Noto Sans CJK JP', 'Hiragino Sans', 'Yu Gothic', 'sans-serif'],
	'MS Mincho': ['Noto Serif CJK JP', 'Hiragino Mincho ProN', 'Yu Mincho', 'serif'],
	SimSun: ['Noto Serif CJK SC', 'STSong', 'serif'],
	SimHei: ['Noto Sans CJK SC', 'STHeiti', 'sans-serif'],
	'Microsoft YaHei': ['Noto Sans CJK SC', 'PingFang SC', 'sans-serif'],
	NSimSun: ['Noto Serif CJK SC', 'STSong', 'serif'],
	FangSong: ['Noto Serif CJK SC', 'STFangsong', 'serif'],
	KaiTi: ['Noto Serif CJK SC', 'STKaiti', 'serif'],
	DengXian: ['Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'sans-serif'],
	等线: ['DengXian', 'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'sans-serif'],
	宋体: ['SimSun', 'Songti SC', 'STSong', 'Noto Serif CJK SC', 'serif'],
	新宋体: ['NSimSun', 'Songti SC', 'STSong', 'Noto Serif CJK SC', 'serif'],
	黑体: ['SimHei', 'Heiti SC', 'STHeiti', 'PingFang SC', 'Noto Sans CJK SC', 'sans-serif'],
	仿宋: ['FangSong', 'STFangsong', 'Songti SC', 'Noto Serif CJK SC', 'serif'],
	楷体: ['KaiTi', 'Kaiti SC', 'STKaiti', 'Noto Serif CJK SC', 'serif'],
	Batang: ['Noto Serif CJK KR', 'AppleMyungjo', 'serif'],
	Dotum: ['Noto Sans CJK KR', 'AppleGothic', 'sans-serif'],
	Gulim: ['Noto Sans CJK KR', 'AppleGothic', 'sans-serif'],
	Malgun: ['Noto Sans CJK KR', 'Apple SD Gothic Neo', 'sans-serif'],
	'Malgun Gothic': ['Noto Sans CJK KR', 'Apple SD Gothic Neo', 'sans-serif'],
	// Source Han Sans (Adobe/Google) ships the same glyphs as Noto Sans CJK;
	// the zh-CN name 思源黑体 is how decks authored in Chinese reference it.
	'思源黑体 CN': [
		'Source Han Sans CN',
		'Noto Sans CJK SC',
		'Noto Sans SC',
		'Microsoft YaHei',
		'sans-serif',
	],
	思源黑体: [
		'Source Han Sans SC',
		'Noto Sans CJK SC',
		'Noto Sans SC',
		'Microsoft YaHei',
		'sans-serif',
	],
	'Source Han Sans CN': [
		'思源黑体 CN',
		'Noto Sans CJK SC',
		'Noto Sans SC',
		'Microsoft YaHei',
		'sans-serif',
	],
	微软雅黑: ['Microsoft YaHei', 'Noto Sans CJK SC', 'PingFang SC', 'sans-serif'],

	// Complex script fonts
	'Arabic Typesetting': ['Noto Naskh Arabic', 'serif'],
	'Simplified Arabic': ['Noto Sans Arabic', 'sans-serif'],
	'Traditional Arabic': ['Noto Naskh Arabic', 'serif'],
	Mangal: ['Noto Sans Devanagari', 'sans-serif'],
	Vrinda: ['Noto Sans Bengali', 'sans-serif'],
	Raavi: ['Noto Sans Gurmukhi', 'sans-serif'],
	Shruti: ['Noto Sans Gujarati', 'sans-serif'],
	Tunga: ['Noto Sans Kannada', 'sans-serif'],
	Kartika: ['Noto Sans Malayalam', 'sans-serif'],
	Iskoola: ['Noto Sans Sinhala', 'sans-serif'],
	'Iskoola Pota': ['Noto Sans Sinhala', 'sans-serif'],
	Leelawadee: ['Noto Sans Thai', 'sans-serif'],
	'Leelawadee UI': ['Noto Sans Thai', 'sans-serif'],
	'Cordia New': ['Noto Sans Thai', 'sans-serif'],
	DokChampa: ['Noto Sans Lao', 'sans-serif'],
	Nyala: ['Noto Sans Ethiopic', 'sans-serif'],
	MoolBoran: ['Noto Sans Khmer', 'sans-serif'],

	// Decorative / Display
	// Condensed display faces: decks often reference the FULL style name
	// ("Bebas Neue Regular"); the suffix-stripped base resolves here. The
	// fallbacks must stay condensed, or short numeric labels sized to the
	// narrow face overflow their text boxes and wrap (issue #132, "77%").
	'Bebas Neue': ['Oswald', 'Arial Narrow', 'Impact', 'sans-serif'],
	// Office stores named Source Sans weights as separate family names (e.g. a
	// theme authored with "Source Sans Pro ExtraLight" as its own face), while
	// Google Fonts and most local installs expose the base family with a
	// weight axis instead. Keep the authored face first for a browser that
	// does have it installed under that exact name, then fall back to the
	// metrically compatible base family before generic sans-serif.
	'Source Sans Pro ExtraLight': [
		'Source Sans Pro',
		'Calibri Light',
		'Segoe UI Light',
		'Arial',
		'sans-serif',
	],
	'Century Gothic': ['URW Gothic', 'Futura', 'sans-serif'],
	Oswald: ['Agency FB', 'Arial Narrow', 'sans-serif'],
	'Franklin Gothic': ['Liberation Sans', 'Helvetica Neue', 'sans-serif'],
	'Franklin Gothic Medium': ['Liberation Sans', 'Helvetica Neue', 'sans-serif'],
	Garamond: ['EB Garamond', 'Cormorant Garamond', 'serif'],
	'Tw Cen MT': ['Century Gothic', 'Futura', 'sans-serif'],
	Rockwell: ['Roboto Slab', 'Rockwell', 'serif'],
	Candara: ['Liberation Sans', 'Optima', 'sans-serif'],
	Constantia: ['Liberation Serif', 'Palatino', 'serif'],
	Corbel: ['Liberation Sans', 'Lucida Grande', 'sans-serif'],
};
