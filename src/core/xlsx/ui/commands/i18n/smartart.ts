// SmartArt strings: the shared SmartArt UI strings of `ooxml-core/diagram` (the ones Word shows;
// their dotted catalogue keys go through `locales/<lang>/smartart.ts`, since English is not the
// identity there) and Excel's own tab label and notices: English key -> [fr, de, es, zh-CN].
import { DIAGRAM_UI_STRINGS } from '../../../../diagram/strings/index';
import { SMART_ART_NOTICES } from '../../../smart-art-notices';
import type { Translations } from './types';

const SHARED: Translations = Object.fromEntries(
	Object.keys(DIAGRAM_UI_STRINGS.en)
		.filter((key) => !key.includes('.'))
		.map((key) => {
			const of = (strings: Record<string, string>) => strings[key] ?? key;
			return [
				key,
				[
					of(DIAGRAM_UI_STRINGS.fr),
					of(DIAGRAM_UI_STRINGS.de),
					of(DIAGRAM_UI_STRINGS.es),
					of(DIAGRAM_UI_STRINGS['zh-CN']),
				] as const,
			];
		}),
);

export const SMARTART_STRINGS: Translations = {
	...SHARED,
	'SmartArt Design': [
		'Création SmartArt',
		'SmartArt-Design',
		'Diseño de SmartArt',
		'SmartArt 设计',
	],
	'Drawn with an approximate layout computed here; not Excel-identical.': [
		'Dessiné avec une disposition approximative calculée ici ; différent d’Excel.',
		'Mit einem hier berechneten Näherungslayout gezeichnet; nicht identisch mit Excel.',
		'Dibujado con un diseño aproximado calculado aquí; no es idéntico a Excel.',
		'使用此处计算的近似版式绘制；与 Excel 不完全相同。',
	],
	[SMART_ART_NOTICES.cached]: [
		'Affiché à partir du dessin mis en cache dans le fichier (la disposition enregistrée en dernier par Excel) ; la disposition Excel n’est pas recalculée ici. Sans modification, le graphique est conservé tel quel à l’enregistrement ; modifier ses formes réécrit la partie de données et supprime le dessin en cache, qu’Excel recalcule à l’ouverture du fichier.',
		'Aus der im Datei zwischengespeicherten Zeichnung angezeigt (das zuletzt von Excel gespeicherte Layout); das Excel-Layout wird hier nicht neu berechnet. Unverändert bleibt die Grafik beim Speichern erhalten; das Bearbeiten ihrer Formen schreibt den Datenteil neu und verwirft die zwischengespeicherte Zeichnung, die Excel beim Öffnen der Datei neu anordnet.',
		'Se muestra a partir del dibujo almacenado en caché en el archivo (el diseño que Excel guardó por última vez); el diseño de Excel no se vuelve a calcular aquí. Sin cambios, el gráfico se conserva al guardar; editar sus formas reescribe la parte de datos y descarta el dibujo en caché, que Excel vuelve a diseñar al abrir el archivo.',
		'根据文件中缓存的绘图显示（Excel 最后保存的版式）；此处不会重新计算 Excel 版式。未编辑时，保存时图形保持不变；编辑其形状会重写数据部件并删除缓存的绘图，Excel 打开文件时会重新排列。',
	],
	[SMART_ART_NOTICES.placeholder]: [
		'Ce fichier ne contient aucun dessin en cache utilisable pour le graphique SmartArt ; il est donc dessiné avec une disposition approximative calculée à partir de ses nœuds, et non avec celle d’Excel. Excel recalcule la disposition à l’ouverture du fichier.',
		'Diese Datei enthält keine verwendbare zwischengespeicherte Zeichnung für die SmartArt-Grafik; sie wird daher mit einem aus ihren Knoten berechneten Näherungslayout gezeichnet, nicht mit dem Layout von Excel. Excel ordnet die Grafik beim Öffnen der Datei neu an.',
		'Este archivo no tiene un dibujo en caché utilizable para el gráfico SmartArt, por lo que se dibuja con un diseño aproximado calculado a partir de sus nodos, no con el diseño de Excel. Excel vuelve a diseñar el gráfico al abrir el archivo.',
		'此文件没有可用的 SmartArt 图形缓存绘图，因此使用根据其节点计算的近似版式绘制，而不是 Excel 计算的版式。Excel 打开文件时会重新排列该图形。',
	],
	[SMART_ART_NOTICES.edited]: [
		'Ce graphique a été modifié ou inséré ici ; il est donc dessiné avec une disposition approximative calculée à partir de ses nœuds, et non avec celle d’Excel. Excel recalcule la disposition à partir de ses parties de données et de disposition à l’ouverture du fichier.',
		'Diese Grafik wurde hier bearbeitet oder eingefügt; sie wird daher mit einem aus ihren Knoten berechneten Näherungslayout gezeichnet, nicht mit dem Layout von Excel. Excel ordnet sie beim Öffnen der Datei aus ihren Daten- und Layoutteilen neu an.',
		'Este gráfico se editó o insertó aquí, por lo que se dibuja con un diseño aproximado calculado a partir de sus nodos, no con el diseño de Excel. Excel vuelve a diseñarlo a partir de sus partes de datos y de diseño al abrir el archivo.',
		'此图形已在此处编辑或插入，因此使用根据其节点计算的近似版式绘制，而不是 Excel 计算的版式。Excel 打开文件时会根据其数据和版式部件重新排列。',
	],
};
