/**
 * Display strings of the shared ink tools every product's Draw tab uses: the Ink Equation dialog
 * (Ink to Math, Ink Equation) and the ruler stencil. Keys are the English text; a product passes
 * its locale and the shared UI translates itself, so the strings are in every locale at once.
 */
import { normalizeLocale } from './locale';

const ROWS = [
	// [en, fr, de, es, zh-CN, ja, ko]
	[
		'Ink Equation',
		'Équation manuscrite',
		'Freihandgleichung',
		'Ecuación manuscrita',
		'墨迹公式',
		'インク数式',
		'잉크 수식',
	],
	[
		'Ink to Math',
		'Convertir l’encre en équation',
		'Freihand in Formel',
		'Entrada de lápiz a matemáticas',
		'将墨迹转换为数学公式',
		'インクを数式に変換',
		'잉크를 수식으로 변환',
	],
	[
		'Write math here',
		'Écrivez les mathématiques ici',
		'Hier Formel schreiben',
		'Escriba la expresión matemática aquí',
		'在此处书写数学公式',
		'ここに数式を書きます',
		'여기에 수식을 쓰세요',
	],
	['Clear', 'Effacer', 'Löschen', 'Borrar', '清除', 'クリア', '지우기'],
	[
		'Recognizing...',
		'Reconnaissance...',
		'Wird erkannt...',
		'Reconociendo...',
		'正在识别...',
		'認識しています...',
		'인식하는 중...',
	],
	[
		'The handwriting could not be recognized. Write it again or correct the LaTeX.',
		'L’écriture n’a pas pu être reconnue. Écrivez-la de nouveau ou corrigez le LaTeX.',
		'Die Handschrift konnte nicht erkannt werden. Schreiben Sie sie erneut, oder korrigieren Sie den LaTeX-Code.',
		'No se pudo reconocer la escritura. Escríbala de nuevo o corrija el LaTeX.',
		'无法识别手写内容。请重新书写或更正 LaTeX。',
		'手書きを認識できませんでした。もう一度書くか、LaTeX を修正してください。',
		'필기를 인식할 수 없습니다. 다시 쓰거나 LaTeX를 수정하세요.',
	],
	[
		'LaTeX equation',
		'Équation LaTeX',
		'LaTeX-Formel',
		'Ecuación LaTeX',
		'LaTeX 公式',
		'LaTeX 数式',
		'LaTeX 수식',
	],
	['Equation', 'Équation', 'Formel', 'Ecuación', '公式', '数式', '수식'],
	['Preview', 'Aperçu', 'Vorschau', 'Vista previa', '预览', 'プレビュー', '미리 보기'],
	[
		'Write an equation to see a preview.',
		'Écrivez une équation pour afficher un aperçu.',
		'Schreiben Sie eine Formel, um eine Vorschau anzuzeigen.',
		'Escriba una ecuación para ver una vista previa.',
		'书写公式以查看预览。',
		'数式を書くとプレビューが表示されます。',
		'수식을 쓰면 미리 보기가 표시됩니다.',
	],
	[
		'This LaTeX could not be converted to an equation.',
		'Ce LaTeX n’a pas pu être converti en équation.',
		'Dieser LaTeX-Code konnte nicht in eine Formel konvertiert werden.',
		'Este LaTeX no se pudo convertir en una ecuación.',
		'无法将此 LaTeX 转换为公式。',
		'この LaTeX は数式に変換できませんでした。',
		'이 LaTeX를 수식으로 변환할 수 없습니다.',
	],
	['Insert', 'Insérer', 'Einfügen', 'Insertar', '插入', '挿入', '삽입'],
	['Cancel', 'Annuler', 'Abbrechen', 'Cancelar', '取消', 'キャンセル', '취소'],
	['Close', 'Fermer', 'Schließen', 'Cerrar', '关闭', '閉じる', '닫기'],
	['Ruler', 'Règle', 'Lineal', 'Regla', '标尺', 'ルーラー', '눈금자'],
	[
		'Drag or use the arrow keys to move. Hold Shift and press an arrow key, or turn the mouse wheel, to rotate.',
		'Faites glisser ou utilisez les touches de direction pour déplacer. Maintenez Maj et appuyez sur une touche de direction, ou tournez la molette, pour faire pivoter.',
		'Ziehen oder die Pfeiltasten verwenden, um zu verschieben. Zum Drehen UMSCHALT gedrückt halten und eine Pfeiltaste drücken oder das Mausrad drehen.',
		'Arrastre o use las teclas de dirección para mover. Mantenga presionada Mayús y presione una tecla de dirección, o gire la rueda del mouse, para girar.',
		'拖动或使用箭头键移动。按住 Shift 并按箭头键或转动鼠标滚轮以旋转。',
		'ドラッグするか方向キーで移動します。回転するには、Shift キーを押しながら方向キーを押すか、マウス ホイールを回します。',
		'끌거나 화살표 키를 사용하여 이동합니다. 회전하려면 Shift 키를 누른 채 화살표 키를 누르거나 마우스 휠을 돌리세요.',
	],
	[
		'Show the ruler and draw straight lines along its edge.',
		'Afficher la règle et tracer des lignes droites le long de son bord.',
		'Das Lineal anzeigen und gerade Linien an seiner Kante zeichnen.',
		'Mostrar la regla y dibujar líneas rectas a lo largo de su borde.',
		'显示标尺并沿其边缘绘制直线。',
		'ルーラーを表示し、その端に沿って直線を描きます。',
		'눈금자를 표시하고 가장자리를 따라 직선을 그립니다.',
	],
	[
		'{angle} degrees',
		'{angle} degrés',
		'{angle} Grad',
		'{angle} grados',
		'{angle} 度',
		'{angle} 度',
		'{angle}도',
	],
] as const satisfies readonly (readonly [string, string, string, string, string, string, string])[];

const LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN', 'ja', 'ko'] as const;

/** One of the shared ink tools' strings (its English text). */
export type InkText = (typeof ROWS)[number][0];

const tables = new Map<string, ReadonlyMap<string, string>>();

/**
 * `text` in `locale`. Any BCP 47 tag resolves as the product UIs resolve theirs
 * (`normalizeLocale`); an unsupported one reads English. `{angle}` is filled from `values`.
 */
export function translateInkText(
	text: InkText,
	locale: string | undefined,
	values: Readonly<Record<string, string | number>> = {},
): string {
	const code = normalizeLocale(locale, LOCALES, 'en');
	let table = tables.get(code);
	if (!table) {
		const column = LOCALES.indexOf(code);
		table = new Map(ROWS.map((row) => [row[0], row[column] ?? row[0]]));
		tables.set(code, table);
	}
	return (table.get(text) ?? text).replace(/\{(\w+)\}/g, (match, key: string) =>
		key in values ? String(values[key]) : match,
	);
}
