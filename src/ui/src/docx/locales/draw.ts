// The Draw tab: tools, pens, conversions, Ink Replay and their disabled reasons. Loaded with the
// tab (`draw-tab.ts` merges them into the display locales), so an editor that never draws never
// downloads them; the tab's own name is in the ribbon tables.
const entries = [
	['Draw', 'Dessin', 'Zeichnen', 'Dibujar', '绘图'],
	['Drawing Tools', 'Outils de dessin', 'Zeichentools', 'Herramientas de dibujo', '绘图工具'],
	['Touch', 'Tactile', 'Toucheingabe', 'Táctil', '触摸'],
	['Convert', 'Convertir', 'Konvertieren', 'Convertir', '转换'],
	['Replay', 'Relecture', 'Wiedergabe', 'Reproducción', '重播'],
	['Lasso Select', 'Sélection par lasso', 'Lassoauswahl', 'Selección de lazo', '套索选择'],
	['Thickness', 'Épaisseur', 'Stärke', 'Grosor', '粗细'],
	['Add Pen', 'Ajouter un stylet', 'Stift hinzufügen', 'Agregar lápiz', '添加笔'],
	['Pen', 'Stylet', 'Stift', 'Lápiz', '笔'],
	['Pencil', 'Crayon', 'Bleistift', 'Lápiz de grafito', '铅笔'],
	['Highlighter', 'Surligneur', 'Textmarker', 'Marcador de resaltado', '荧光笔'],
	['Action Pen', 'Stylet d’action', 'Aktionsstift', 'Lápiz de acción', '操作笔'],
	['Pens', 'Stylets', 'Stifte', 'Lápices', '笔'],
	['Delete Pen', 'Supprimer le stylet', 'Stift löschen', 'Eliminar lápiz', '删除笔'],
	[
		'Draw with Touch',
		'Dessiner avec la fonction tactile',
		'Mit Toucheingabe zeichnen',
		'Dibujar con función táctil',
		'用触摸进行绘图',
	],
	[
		'Ink to Shape',
		'Convertir l’encre en forme',
		'Freihand in Form',
		'Entrada de lápiz a forma',
		'将墨迹转换为形状',
	],
	[
		'Ink to Math',
		'Convertir l’encre en équation',
		'Freihand in Formel',
		'Entrada de lápiz a matemáticas',
		'将墨迹转换为数学公式',
	],
	[
		'Ink Replay',
		'Relecture de l’encre',
		'Freihandwiedergabe',
		'Reproducción de entrada de lápiz',
		'墨迹重播',
	],
	['Highlight Yellow', 'Jaune surligneur', 'Textmarkergelb', 'Amarillo de resaltado', '荧光黄'],
	['Turquoise', 'Turquoise', 'Türkis', 'Turquesa', '青绿色'],
	['Pink', 'Rose', 'Rosa', 'Rosa', '粉红色'],
	['Gray', 'Gris', 'Grau', 'Gris', '灰色'],
	['Ink', 'Encre', 'Freihand', 'Entrada de lápiz', '墨迹'],
	['Drawing canvas', 'Zone de dessin', 'Zeichenbereich', 'Lienzo de dibujo', '绘图画布'],
	[
		'Select ink with Lasso Select or select an ink drawing.',
		'Sélectionnez de l’encre avec la sélection par lasso ou sélectionnez un dessin à l’encre.',
		'Wählen Sie Freihand mit der Lassoauswahl oder eine Freihandzeichnung aus.',
		'Seleccione entrada de lápiz con Selección de lazo o seleccione un dibujo de entrada de lápiz.',
		'请用套索选择墨迹，或选择一个墨迹绘图。',
	],
	[
		'Ink to Math needs a handwriting recognizer, which this editor does not include.',
		'Convertir l’encre en équation nécessite un module de reconnaissance de l’écriture manuscrite, que cet éditeur n’inclut pas.',
		'Freihand in Formel benötigt eine Handschrifterkennung, die dieser Editor nicht enthält.',
		'Entrada de lápiz a matemáticas necesita un reconocedor de escritura a mano que este editor no incluye.',
		'将墨迹转换为数学公式需要手写识别器，此编辑器未包含。',
	],
	[
		'There is no ink on the pages to replay.',
		'Il n’y a pas d’encre à relire sur les pages.',
		'Auf den Seiten ist keine Freihand zum Wiedergeben vorhanden.',
		'No hay entrada de lápiz en las páginas para reproducir.',
		'页面上没有可重播的墨迹。',
	],
	[
		'The pen gallery is full.',
		'La galerie de stylets est pleine.',
		'Der Stiftkatalog ist voll.',
		'La galería de lápices está llena.',
		'笔库已满。',
	],
	['Stencils', 'Gabarits', 'Schablonen', 'Plantillas', '模具'],
	[
		'The ruler is not available in this editor.',
		'La règle n’est pas disponible dans cet éditeur.',
		'Das Lineal ist in diesem Editor nicht verfügbar.',
		'La regla no está disponible en este editor.',
		'此编辑器中没有标尺。',
	],
	[
		'Draw ink with this pen. Select it again to change its color and thickness.',
		'Dessinez à l’encre avec ce stylet. Sélectionnez-le de nouveau pour modifier sa couleur et son épaisseur.',
		'Mit diesem Stift freihand zeichnen. Wählen Sie ihn erneut aus, um Farbe und Stärke zu ändern.',
		'Dibuje con este lápiz. Vuelva a seleccionarlo para cambiar su color y grosor.',
		'用此笔绘制墨迹。再次选择它可更改颜色和粗细。',
	],
	[
		'Select objects and text instead of drawing.',
		'Sélectionnez des objets et du texte au lieu de dessiner.',
		'Objekte und Text auswählen statt zu zeichnen.',
		'Seleccione objetos y texto en lugar de dibujar.',
		'选择对象和文本，而不是绘图。',
	],
	[
		'Draw around ink to select it.',
		'Entourez l’encre pour la sélectionner.',
		'Zeichnen Sie um Freihand herum, um sie auszuwählen.',
		'Dibuje alrededor de la entrada de lápiz para seleccionarla.',
		'围绕墨迹绘制以将其选中。',
	],
	[
		'Erase the ink strokes you touch.',
		'Effacez les traits d’encre que vous touchez.',
		'Die berührten Freihandstriche löschen.',
		'Borre los trazos de entrada de lápiz que toque.',
		'擦除所触及的墨迹笔画。',
	],
	[
		'Draw with your finger as well as a pen or mouse.',
		'Dessinez avec le doigt, ainsi qu’avec un stylet ou une souris.',
		'Mit dem Finger sowie mit Stift oder Maus zeichnen.',
		'Dibuje con el dedo, además de con un lápiz o un mouse.',
		'除了笔或鼠标，还可以用手指绘图。',
	],
	[
		'Turn ink you drew into a shape.',
		'Transformez l’encre que vous avez dessinée en forme.',
		'Gezeichnete Freihand in eine Form umwandeln.',
		'Convierta la entrada de lápiz dibujada en una forma.',
		'将绘制的墨迹转换为形状。',
	],
	[
		'Turn handwritten math into an equation.',
		'Transformez des mathématiques manuscrites en équation.',
		'Handschriftliche Mathematik in eine Formel umwandeln.',
		'Convierta matemáticas escritas a mano en una ecuación.',
		'将手写数学转换为公式。',
	],
	[
		'Watch the ink being drawn again, in the order it was drawn.',
		'Regardez l’encre se dessiner de nouveau, dans l’ordre où elle a été tracée.',
		'Die Freihand in der gezeichneten Reihenfolge erneut entstehen sehen.',
		'Vea cómo se vuelve a dibujar la entrada de lápiz, en el orden en que se dibujó.',
		'按绘制顺序重新观看墨迹的绘制过程。',
	],
	['More Colors...', 'Plus de couleurs...', 'Weitere Farben...', 'Más colores...', '其他颜色...'],
	// Word's own table eraser wording, kept the same (that table loads with the Table tabs only).
	['Eraser', 'Gomme', 'Radierer', 'Borrador', '擦除'],
	[
		'Editing is unavailable.',
		'La modification est indisponible.',
		'Die Bearbeitung ist nicht verfügbar.',
		'La edición no está disponible.',
		'无法编辑。',
	],
	[
		'Ink can only be drawn on a page with text.',
		'L’encre ne peut être dessinée que sur une page contenant du texte.',
		'Freihand kann nur auf einer Seite mit Text gezeichnet werden.',
		'La entrada de lápiz solo se puede dibujar en una página con texto.',
		'只能在有文字的页面上绘制墨迹。',
	],
	[
		'Ink to Shape did not recognize a shape in the selected ink.',
		'Convertir l’encre en forme n’a reconnu aucune forme dans l’encre sélectionnée.',
		'Freihand in Form hat in der ausgewählten Freihand keine Form erkannt.',
		'Entrada de lápiz a forma no reconoció ninguna forma en la entrada de lápiz seleccionada.',
		'“将墨迹转换为形状”未在所选墨迹中识别出形状。',
	],
] as const;

/** A Draw tab string (English text). */
export type DrawKey = (typeof entries)[number][0];
type Key = DrawKey;
const dictionary = (col: 0 | 1 | 2 | 3 | 4): Record<Key, string> =>
	Object.fromEntries(entries.map((row) => [row[0], row[col]])) as Record<Key, string>;

export const enDraw = dictionary(0);
export const frDraw = dictionary(1);
export const deDraw = dictionary(2);
export const esDraw = dictionary(3);
export const zhDraw = dictionary(4);
