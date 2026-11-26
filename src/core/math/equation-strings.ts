/**
 * The Equation tab's display strings in every supported locale: its commands, the structure
 * gallery names and the built-in equations' names (the gallery entries are in
 * `equation-gallery-strings.ts`). Keys are the English text the editors
 * look up; each row is `[English, French, German, Spanish, Simplified Chinese]`.
 */
type Row = readonly [string, string, string, string, string];

const ROWS = [
	['Tools', 'Outils', 'Tools', 'Herramientas', '工具'],
	['Structures', 'Structures', 'Strukturen', 'Estructuras', '结构'],
	['Ink Equation', 'Équation manuscrite', 'Freihandgleichung', 'Ecuación manuscrita', '墨迹公式'],
	[
		'Ink Equation needs handwriting recognition, which this editor does not have.',
		"L'équation manuscrite nécessite la reconnaissance de l'écriture, que cet éditeur ne propose pas.",
		'Die Freihandgleichung erfordert eine Handschrifterkennung, die dieser Editor nicht hat.',
		'La ecuación manuscrita requiere reconocimiento de escritura, que este editor no tiene.',
		'墨迹公式需要手写识别，此编辑器不提供该功能。',
	],
	['Convert', 'Convertir', 'Konvertieren', 'Convertir', '转换'],
	[
		'Current - Professional',
		'Actuelle - Professionnel',
		'Aktuell - Professionell',
		'Actual - Profesional',
		'当前 - 专业型',
	],
	['Current - Linear', 'Actuelle - Linéaire', 'Aktuell - Linear', 'Actual - Lineal', '当前 - 线性'],
	[
		'All - Professional',
		'Toutes - Professionnel',
		'Alle - Professionell',
		'Todas - Profesional',
		'全部 - 专业型',
	],
	['All - Linear', 'Toutes - Linéaire', 'Alle - Linear', 'Todas - Lineal', '全部 - 线性'],
	['Normal Text', 'Texte normal', 'Normaler Text', 'Texto normal', '普通文本'],
	['Unicode', 'Unicode', 'Unicode', 'Unicode', 'Unicode'],
	['LaTeX', 'LaTeX', 'LaTeX', 'LaTeX', 'LaTeX'],
	[
		'Equation input format',
		"Format d'entrée des équations",
		'Eingabeformat für Formeln',
		'Formato de entrada de ecuaciones',
		'公式输入格式',
	],
	[
		'Insert New Equation',
		'Insérer une nouvelle équation',
		'Neue Formel einfügen',
		'Insertar nueva ecuación',
		'插入新公式',
	],
	['Equation editor', "Éditeur d'équations", 'Formeleditor', 'Editor de ecuaciones', '公式编辑器'],
	[
		'Type equation here.',
		'Tapez une équation ici.',
		'Geben Sie hier eine Formel ein.',
		'Escriba aquí una ecuación.',
		'在此处键入公式。',
	],
	[
		'Symbol category',
		'Catégorie de symboles',
		'Symbolkategorie',
		'Categoría de símbolos',
		'符号类别',
	],
	// Built-in equations (the four first ones are the Insert menu's own strings).
	[
		'Expansion of a Sum',
		"Développement d'une somme",
		'Entwicklung einer Summe',
		'Desarrollo de una suma',
		'和的展开式',
	],
	['Fourier Series', 'Série de Fourier', 'Fourier-Reihe', 'Serie de Fourier', '傅立叶级数'],
	[
		'Taylor Expansion',
		'Développement de Taylor',
		'Taylor-Entwicklung',
		'Desarrollo de Taylor',
		'泰勒展开式',
	],
	[
		'Trig Identity 1',
		'Identité trigonométrique 1',
		'Trigonometrische Identität 1',
		'Identidad trigonométrica 1',
		'三角恒等式 1',
	],
	[
		'Trig Identity 2',
		'Identité trigonométrique 2',
		'Trigonometrische Identität 2',
		'Identidad trigonométrica 2',
		'三角恒等式 2',
	],
	// Structure galleries.
	['Fraction', 'Fraction', 'Bruch', 'Fracción', '分式'],
	['Script', 'Script', 'Skript', 'Script', '上下标'],
	['Radical', 'Radical', 'Wurzel', 'Radical', '根式'],
	['Integral', 'Intégrale', 'Integral', 'Integral', '积分'],
	['Large Operator', 'Grand opérateur', 'Großer Operator', 'Operador grande', '大型运算符'],
	['Bracket', 'Crochet', 'Klammer', 'Corchete', '括号'],
	['Function', 'Fonction', 'Funktion', 'Función', '函数'],
	['Accent', 'Accent', 'Akzent', 'Acento', '导数符号'],
	[
		'Limit and Log',
		'Limite et log',
		'Grenzwert und Logarithmus',
		'Límite y logaritmo',
		'极限和对数',
	],
	['Operator', 'Opérateur', 'Operator', 'Operador', '运算符'],
	['Matrix', 'Matrice', 'Matrix', 'Matriz', '矩阵'],
] as const satisfies readonly Row[];

type EquationKey = (typeof ROWS)[number][0];
export type EquationUiStrings = Record<EquationKey, string>;

const column = (index: 0 | 1 | 2 | 3 | 4) =>
	Object.fromEntries(ROWS.map((row) => [row[0], row[index]])) as EquationUiStrings;

/** The Equation tab's strings by locale, for the products' dictionaries. */
export const EQUATION_UI_STRINGS: Readonly<
	Record<'en' | 'fr' | 'de' | 'es' | 'zh-CN', EquationUiStrings>
> = {
	en: column(0),
	fr: column(1),
	de: column(2),
	es: column(3),
	'zh-CN': column(4),
};
