import type { Translations } from '../types';

/** Insert Function descriptions of the cube, host-provider and width-conversion functions. */
export const FUNCTION_HOST_CUBE_STRINGS: Translations = {
	'Converts full-width (double-byte) letters and Katakana to half-width (single-byte) characters.':
		[
			'Convertit les lettres et les katakana pleine largeur (deux octets) en caractères demi-largeur (un octet).',
			'Wandelt Buchstaben und Katakana mit voller Breite (Doppelbyte) in Zeichen mit halber Breite (Einzelbyte) um.',
			'Convierte letras y katakana de ancho completo (dos bytes) en caracteres de ancho medio (un byte).',
			'将全角(双字节)字母和片假名转换为半角(单字节)字符。',
		],
	'Converts half-width (single-byte) letters and Katakana to full-width (double-byte) characters.':
		[
			'Convertit les lettres et les katakana demi-largeur (un octet) en caractères pleine largeur (deux octets).',
			'Wandelt Buchstaben und Katakana mit halber Breite (Einzelbyte) in Zeichen mit voller Breite (Doppelbyte) um.',
			'Convierte letras y katakana de ancho medio (un byte) en caracteres de ancho completo (dos bytes).',
			'将半角(单字节)字母和片假名转换为全角(双字节)字符。',
		],
	'Calls a host-supplied AI provider with prompt text and optional cell or range context.': [
		"Appelle un fournisseur d'IA fourni par l'hôte avec un texte d'invite et un contexte facultatif de cellule ou de plage.",
		'Ruft einen vom Host bereitgestellten KI-Anbieter mit einem Prompttext und optionalem Zellen- oder Bereichskontext auf.',
		'Llama a un proveedor de IA suministrado por el host con texto de solicitud y contexto opcional de celda o rango.',
		'使用提示文本和可选的单元格或区域上下文调用主机提供的 AI 服务。',
	],
	'Returns a key performance indicator (KPI) property from a cube.': [
		"Renvoie une propriété d'indicateur de performance clé (KPI) d'un cube.",
		'Gibt eine Eigenschaft eines Key Performance Indicators (KPI) aus einem Cube zurück.',
		'Devuelve una propiedad de indicador clave de rendimiento (KPI) de un cubo.',
		'从多维数据集返回关键绩效指标(KPI)属性。',
	],
	'Returns a member or tuple from a cube.': [
		"Renvoie un membre ou un tuple d'un cube.",
		'Gibt ein Element oder Tupel aus einem Cube zurück.',
		'Devuelve un miembro o una tupla de un cubo.',
		'从多维数据集返回成员或元组。',
	],
	'Returns the value of a member property from a cube.': [
		"Renvoie la valeur d'une propriété de membre d'un cube.",
		'Gibt den Wert einer Elementeigenschaft aus einem Cube zurück.',
		'Devuelve el valor de una propiedad de miembro de un cubo.',
		'从多维数据集返回成员属性的值。',
	],
	'Returns the nth, or ranked, member in a set.': [
		"Renvoie le n-ième membre, ou le membre classé, d'un ensemble.",
		'Gibt das n-te bzw. rangierte Element in einer Menge zurück.',
		'Devuelve el miembro n-ésimo, o clasificado, de un conjunto.',
		'返回集合中的第 n 个(或按排名的)成员。',
	],
	'Defines a calculated set of members or tuples from a cube.': [
		"Définit un ensemble calculé de membres ou de tuples d'un cube.",
		'Definiert eine berechnete Menge von Elementen oder Tupeln aus einem Cube.',
		'Define un conjunto calculado de miembros o tuplas de un cubo.',
		'定义来自多维数据集的成员或元组的计算集。',
	],
	'Returns the number of items in a set.': [
		"Renvoie le nombre d'éléments d'un ensemble.",
		'Gibt die Anzahl der Elemente in einer Menge zurück.',
		'Devuelve el número de elementos de un conjunto.',
		'返回集合中的项数。',
	],
	'Returns an aggregated value from a cube.': [
		"Renvoie une valeur agrégée d'un cube.",
		'Gibt einen aggregierten Wert aus einem Cube zurück.',
		'Devuelve un valor agregado de un cubo.',
		'从多维数据集返回汇总值。',
	],
	'Identifies the language of text using a host-supplied provider.': [
		"Identifie la langue d'un texte à l'aide d'un fournisseur fourni par l'hôte.",
		'Erkennt die Sprache eines Textes mithilfe eines vom Host bereitgestellten Anbieters.',
		'Identifica el idioma de un texto mediante un proveedor suministrado por el host.',
		'使用主机提供的服务识别文本的语言。',
	],
	'Converts a number to euros, from euros to a euro member currency, or between euro member currencies using fixed EU rates.':
		[
			"Convertit un nombre en euros, des euros vers une devise d'un pays membre de la zone euro, ou entre devises de la zone euro, aux taux fixes de l'UE.",
			'Rechnet eine Zahl in Euro, von Euro in eine Währung eines Euro-Mitgliedstaats oder zwischen diesen Währungen zu den festen EU-Kursen um.',
			'Convierte un número a euros, de euros a una moneda de un país miembro del euro, o entre monedas de países miembros, con los tipos fijos de la UE.',
			'使用欧盟固定汇率将数字转换为欧元、从欧元转换为欧元区成员国货币,或在欧元区成员国货币之间转换。',
		],
	'Resolves an image URL through a host-supplied provider.': [
		"Résout l'URL d'une image par l'intermédiaire d'un fournisseur fourni par l'hôte.",
		'Löst eine Bild-URL über einen vom Host bereitgestellten Anbieter auf.',
		'Resuelve la URL de una imagen mediante un proveedor suministrado por el host.',
		'通过主机提供的服务解析图像 URL。',
	],
	'Information about the current workbook and calculation environment.': [
		"Informations sur le classeur actif et l'environnement de calcul.",
		'Informationen über die aktuelle Arbeitsmappe und die Berechnungsumgebung.',
		'Información sobre el libro actual y el entorno de cálculo.',
		'有关当前工作簿和计算环境的信息。',
	],
	'Extracts the phonetic (furigana) characters from a text string reference.': [
		"Extrait les caractères phonétiques (furigana) d'une référence de chaîne de texte.",
		'Extrahiert die phonetischen Zeichen (Furigana) aus einem Textzeichenfolgenbezug.',
		'Extrae los caracteres fonéticos (furigana) de una referencia de cadena de texto.',
		'从文本字符串引用中提取拼音(注音假名)字符。',
	],
	'Retrieves real-time data from a host-supplied provider.': [
		"Récupère des données en temps réel auprès d'un fournisseur fourni par l'hôte.",
		'Ruft Echtzeitdaten von einem vom Host bereitgestellten Anbieter ab.',
		'Recupera datos en tiempo real de un proveedor suministrado por el host.',
		'从主机提供的服务检索实时数据。',
	],
	'Retrieves historical market data from a host-supplied provider.': [
		"Récupère des données de marché historiques auprès d'un fournisseur fourni par l'hôte.",
		'Ruft historische Marktdaten von einem vom Host bereitgestellten Anbieter ab.',
		'Recupera datos históricos del mercado de un proveedor suministrado por el host.',
		'从主机提供的服务检索历史市场数据。',
	],
	'Translates text using a host-supplied translation provider.': [
		"Traduit un texte à l'aide d'un fournisseur de traduction fourni par l'hôte.",
		'Übersetzt Text mithilfe eines vom Host bereitgestellten Übersetzungsanbieters.',
		'Traduce texto mediante un proveedor de traducción suministrado por el host.',
		'使用主机提供的翻译服务翻译文本。',
	],
	'Retrieves data from an HTTP(S) URL through a host-supplied provider.': [
		"Récupère des données d'une URL HTTP(S) par l'intermédiaire d'un fournisseur fourni par l'hôte.",
		'Ruft Daten von einer HTTP(S)-URL über einen vom Host bereitgestellten Anbieter ab.',
		'Recupera datos de una URL HTTP(S) mediante un proveedor suministrado por el host.',
		'通过主机提供的服务从 HTTP(S) URL 检索数据。',
	],
};
