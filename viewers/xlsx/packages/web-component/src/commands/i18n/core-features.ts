// Strings of the features the core gained in ooxml-core 0.4 (date and time entry, "A Date
// Occurring" rules, the full built-in cell style list, print centring, protection passwords,
// outline axes) and the core's own edit error messages, which the UI shows translated:
// English key -> [fr, de, es, zh-CN].
import type { Translations } from './types.js';

const ACCENT: readonly [fr: string, de: string, es: string, zhCN: string] = [
	'Accent',
	'Akzent',
	'Énfasis',
	'着色 ',
];

/** `40% - Accent1` ... `60% - Accent6`, the way each Excel language spells them. */
const accentShades: Translations = Object.fromEntries(
	[40, 60].flatMap((pct) =>
		[1, 2, 3, 4, 5, 6].map((n) => [
			`${pct}% - Accent${n}`,
			[
				`${pct} % - ${ACCENT[0]}${n}`,
				`${pct} % - ${ACCENT[1]}${n}`,
				`${pct}% - ${ACCENT[2]}${n}`,
				`${pct}% - ${ACCENT[3]}${n}`,
			] as const,
		]),
	),
);

export const CORE_FEATURE_STRINGS: Translations = {
	...accentShades,
	'Comma [0]': ['Milliers [0]', 'Komma [0]', 'Millares [0]', '千位分隔[0]'],
	'Currency [0]': ['Monétaire [0]', 'Währung [0]', 'Moneda [0]', '货币[0]'],
	Hyperlink: ['Lien hypertexte', 'Link', 'Hipervínculo', '超链接'],
	'Followed Hyperlink': [
		'Lien hypertexte visité',
		'Besuchter Link',
		'Hipervínculo visitado',
		'已访问的超链接',
	],
	// Date and time entry (Ctrl+; and Ctrl+Shift+;)
	'Insert the current date': [
		'Insérer la date du jour',
		'Aktuelles Datum einfügen',
		'Insertar la fecha actual',
		'插入当前日期',
	],
	'Insert the current time': [
		"Insérer l'heure actuelle",
		'Aktuelle Uhrzeit einfügen',
		'Insertar la hora actual',
		'插入当前时间',
	],
	// A Date Occurring
	'A Date Occurring...': ['Une date se produisant...', 'Datum...', 'Una fecha...', '发生日期...'],
	'A Date Occurring': ['Une date se produisant', 'Datum', 'Una fecha', '发生日期'],
	'Format cells that contain a date occurring:': [
		'Appliquer une mise en forme aux cellules contenant une date :',
		'Zellen formatieren, die ein Datum enthalten:',
		'Dar formato a las celdas que contienen una fecha:',
		'为包含以下日期的单元格设置格式:',
	],
	'Dates Occurring': ['Dates se produisant', 'Datumswerte', 'Fechas', '发生日期'],
	Yesterday: ['Hier', 'Gestern', 'Ayer', '昨天'],
	Today: ["Aujourd'hui", 'Heute', 'Hoy', '今天'],
	Tomorrow: ['Demain', 'Morgen', 'Mañana', '明天'],
	'In the last 7 days': [
		'Dans les 7 derniers jours',
		'In den letzten 7 Tagen',
		'En los últimos 7 días',
		'最近 7 天',
	],
	'Last week': ['Semaine dernière', 'Letzte Woche', 'Semana pasada', '上周'],
	'This week': ['Cette semaine', 'Diese Woche', 'Esta semana', '本周'],
	'Next week': ['Semaine prochaine', 'Nächste Woche', 'Semana próxima', '下周'],
	'Last month': ['Mois dernier', 'Letzter Monat', 'Mes pasado', '上个月'],
	'This month': ['Ce mois-ci', 'Dieser Monat', 'Este mes', '本月'],
	'Next month': ['Mois prochain', 'Nächster Monat', 'Mes próximo', '下个月'],
	// Protection
	'Confirmation password is not identical.': [
		"Le mot de passe de confirmation n'est pas identique.",
		'Das Bestätigungskennwort ist nicht identisch.',
		'La contraseña de confirmación no es idéntica.',
		'确认密码不一致。',
	],
	'Password:': ['Mot de passe :', 'Kennwort:', 'Contraseña:', '密码:'],
	'Password (optional):': [
		'Mot de passe (facultatif) :',
		'Kennwort (optional):',
		'Contraseña (opcional):',
		'密码(可选):',
	],
	'Unprotect Workbook': [
		'Ôter la protection du classeur',
		'Arbeitsmappenschutz aufheben',
		'Desproteger libro',
		'撤消工作簿保护',
	],
	'Protect Structure and Windows': [
		'Protéger la structure et les fenêtres',
		'Struktur und Fenster schützen',
		'Proteger estructura y ventanas',
		'保护结构和窗口',
	],
	'Protect workbook for': [
		'Protéger le classeur pour',
		'Arbeitsmappe schützen für',
		'Proteger libro para',
		'保护工作簿',
	],
	Structure: ['Structure', 'Struktur', 'Estructura', '结构'],
	'The password you supplied is not correct.': [
		"Le mot de passe que vous avez fourni n'est pas correct.",
		'Das eingegebene Kennwort ist falsch.',
		'La contraseña proporcionada no es correcta.',
		'您提供的密码不正确。',
	],
	// Errors the core's edits report (shown through t()).
	'A conditional format needs at least one range and one rule.': [
		'Une mise en forme conditionnelle nécessite au moins une plage et une règle.',
		'Eine bedingte Formatierung benötigt mindestens einen Bereich und eine Regel.',
		'Un formato condicional necesita al menos un rango y una regla.',
		'条件格式至少需要一个区域和一条规则。',
	],
	'A table needs at least one row.': [
		'Un tableau doit contenir au moins une ligne.',
		'Eine Tabelle benötigt mindestens eine Zeile.',
		'Una tabla necesita al menos una fila.',
		'表格至少需要一行。',
	],
	'Cannot remove duplicates from a range that contains merged cells.': [
		'Impossible de supprimer les doublons d’une plage contenant des cellules fusionnées.',
		'Duplikate können nicht aus einem Bereich mit verbundenen Zellen entfernt werden.',
		'No se pueden quitar duplicados de un rango con celdas combinadas.',
		'无法从包含合并单元格的区域中删除重复项。',
	],
	'Cannot shift cells: a merged range would be split.': [
		'Impossible de décaler les cellules : une plage fusionnée serait scindée.',
		'Zellen können nicht verschoben werden: Ein verbundener Bereich würde geteilt.',
		'No se pueden desplazar las celdas: se dividiría un rango combinado.',
		'无法移动单元格: 合并区域将被拆分。',
	],
	'Cannot shift cells: a table would be split.': [
		'Impossible de décaler les cellules : un tableau serait scindé.',
		'Zellen können nicht verschoben werden: Eine Tabelle würde geteilt.',
		'No se pueden desplazar las celdas: se dividiría una tabla.',
		'无法移动单元格: 表格将被拆分。',
	],
	'Cannot sort a range that contains part of a dynamic array.': [
		'Impossible de trier une plage contenant une partie d’une matrice dynamique.',
		'Ein Bereich mit einem Teil einer dynamischen Matrix kann nicht sortiert werden.',
		'No se puede ordenar un rango que contiene parte de una matriz dinámica.',
		'无法对包含部分动态数组的区域排序。',
	],
	'The new range must overlap the table.': [
		'La nouvelle plage doit chevaucher le tableau.',
		'Der neue Bereich muss die Tabelle überlappen.',
		'El nuevo rango debe superponerse a la tabla.',
		'新区域必须与表格重叠。',
	],
	'The table headers must remain in the same row.': [
		'Les en-têtes du tableau doivent rester sur la même ligne.',
		'Die Tabellenüberschriften müssen in derselben Zeile bleiben.',
		'Los encabezados de la tabla deben permanecer en la misma fila.',
		'表格标题必须保持在同一行。',
	],
	'The table range is too small.': [
		'La plage du tableau est trop petite.',
		'Der Tabellenbereich ist zu klein.',
		'El rango de la tabla es demasiado pequeño.',
		'表格区域太小。',
	],
	'A table name must start with a letter or underscore and contain no spaces.': [
		'Un nom de tableau doit commencer par une lettre ou un trait de soulignement, sans espaces.',
		'Ein Tabellenname muss mit einem Buchstaben oder Unterstrich beginnen und darf keine Leerzeichen enthalten.',
		'Un nombre de tabla debe empezar por una letra o un guion bajo y no contener espacios.',
		'表名称必须以字母或下划线开头，且不能包含空格。',
	],
	'A table name cannot look like a cell reference.': [
		'Un nom de tableau ne peut pas ressembler à une référence de cellule.',
		'Ein Tabellenname darf nicht wie ein Zellbezug aussehen.',
		'Un nombre de tabla no puede parecer una referencia de celda.',
		'表名称不能类似单元格引用。',
	],
	'That name is already used by another table or name.': [
		'Ce nom est déjà utilisé par un autre tableau ou un autre nom.',
		'Dieser Name wird bereits von einer anderen Tabelle oder einem Namen verwendet.',
		'Ese nombre ya lo usa otra tabla u otro nombre.',
		'该名称已被另一个表格或名称使用。',
	],
};
