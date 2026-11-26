import type { VisioLocale } from '../ribbon-locale';

export const VISIO_STENCIL_STRINGS = {
	en: {
		OK: 'OK',
		Cancel: 'Cancel',
		'Close {name}': 'Close {name}',
		'Legacy binary stencils (.vss, .vst) are not supported.':
			'Legacy binary stencils (.vss, .vst) are not supported.',
		'This file has no shapes.': 'This file has no shapes.',
		'This file is too large to open.': 'This file is too large to open.',
		'Missing ZIP end record': 'Missing ZIP end record',
		'Another part uses this master part.': 'Another part uses this master part.',
		'Edit Stencil': 'Edit Stencil',
		'Turn on Edit Stencil to change its masters.': 'Turn on Edit Stencil to change its masters.',
		'Enter a name.': 'Enter a name.',
		'The stencil exceeded the isolated editing time limit.':
			'The stencil exceeded the isolated editing time limit.',
		'The Visio stencil worker failed.': 'The Visio stencil worker failed.',
		'New Stencil': 'New Stencil',
		'Stencil name': 'Stencil name',
		'Untitled Stencil': 'Untitled Stencil',
		'Stencil actions': 'Stencil actions',
		'Save Stencil': 'Save Stencil',
		'Add Selected Shape': 'Add Selected Shape',
		'Master name': 'Master name',
		'New Master': 'New Master',
		'Rename Master': 'Rename Master',
		'Delete Master': 'Delete Master',
		'Edit Master': 'Edit Master',
		'Default shape text': 'Default shape text',
		'My Shapes': 'My Shapes',
		'Add to My Shapes': 'Add to My Shapes',
		'Remove from My Shapes': 'Remove from My Shapes',
		'Select one editable 2D shape in a drawing.': 'Select one editable 2D shape in a drawing.',
		'Open the source drawing to copy this master.': 'Open the source drawing to copy this master.',
		'Select a top-level shape with editable plain text.':
			'Select a top-level shape with editable plain text.',
		'Enter a master name between 1 and 255 characters.':
			'Enter a master name between 1 and 255 characters.',
		'A master already has that name.': 'A master already has that name.',
		'Enter plain text of up to 4096 characters.': 'Enter plain text of up to 4096 characters.',
		'Close another stencil before opening or editing this one.':
			'Close another stencil before opening or editing this one.',
		'Editing requires a .vssx or .vssm stencil.': 'Editing requires a .vssx or .vssm stencil.',
		'This master depends on another master in its original stencil.':
			'This master depends on another master in its original stencil.',
		'Select a standalone shape or copy its original master.':
			'Select a standalone shape or copy its original master.',
		'Select a standalone 2D shape without a layer.':
			'Select a standalone 2D shape without a layer.',
		'The selected shape needs a positive width and height.':
			'The selected shape needs a positive width and height.',
		'The shape and stencil use different Visio namespaces.':
			'The shape and stencil use different Visio namespaces.',
		'The selected shape uses a non-default font.': 'The selected shape uses a non-default font.',
		'Another master entry uses this master part.': 'Another master entry uses this master part.',
	},
	de: {
		OK: 'Bestätigen',
		Cancel: 'Abbrechen',
		'Close {name}': '{name} schließen',
		'Legacy binary stencils (.vss, .vst) are not supported.':
			'Ältere binäre Schablonen (.vss, .vst) werden nicht unterstützt.',
		'This file has no shapes.': 'Diese Datei enthält keine Shapes.',
		'This file is too large to open.': 'Diese Datei ist zu groß zum Öffnen.',
		'Missing ZIP end record': 'Der ZIP-Enddatensatz fehlt',
		'Another part uses this master part.': 'Ein anderer Teil verwendet diesen Masterteil.',
		'Edit Stencil': 'Schablone bearbeiten',
		'Turn on Edit Stencil to change its masters.':
			'Aktivieren Sie die Schablonenbearbeitung, um Masterformen zu ändern.',
		'Enter a name.': 'Geben Sie einen Namen ein.',
		'The stencil exceeded the isolated editing time limit.':
			'Die Schablonenbearbeitung hat das Zeitlimit überschritten.',
		'The Visio stencil worker failed.':
			'Die Hintergrundverarbeitung der Visio-Schablone ist fehlgeschlagen.',
		'New Stencil': 'Neue Schablone',
		'Stencil name': 'Schablonenname',
		'Untitled Stencil': 'Unbenannte Schablone',
		'Stencil actions': 'Schablonenaktionen',
		'Save Stencil': 'Schablone speichern',
		'Add Selected Shape': 'Ausgewählte Form hinzufügen',
		'Master name': 'Mastername',
		'New Master': 'Neue Masterform',
		'Rename Master': 'Masterform umbenennen',
		'Delete Master': 'Masterform löschen',
		'Edit Master': 'Masterform bearbeiten',
		'Default shape text': 'Standardtext der Form',
		'My Shapes': 'Meine Shapes',
		'Add to My Shapes': 'Zu „Meine Shapes“ hinzufügen',
		'Remove from My Shapes': 'Aus „Meine Shapes“ entfernen',
		'Select one editable 2D shape in a drawing.':
			'Wählen Sie eine bearbeitbare 2D-Form in einer Zeichnung aus.',
		'Open the source drawing to copy this master.':
			'Öffnen Sie zum Kopieren der Masterform die Quellzeichnung.',
		'Select a top-level shape with editable plain text.':
			'Wählen Sie eine Form der obersten Ebene mit bearbeitbarem Klartext.',
		'Enter a master name between 1 and 255 characters.':
			'Geben Sie einen Masterformnamen mit 1 bis 255 Zeichen ein.',
		'A master already has that name.': 'Eine Masterform hat bereits diesen Namen.',
		'Enter plain text of up to 4096 characters.':
			'Geben Sie Klartext mit höchstens 4096 Zeichen ein.',
		'Close another stencil before opening or editing this one.':
			'Schließen Sie eine andere Schablone, bevor Sie diese öffnen oder bearbeiten.',
		'Editing requires a .vssx or .vssm stencil.':
			'Zum Bearbeiten ist eine .vssx- oder .vssm-Schablone erforderlich.',
		'This master depends on another master in its original stencil.':
			'Diese Masterform hängt von einer anderen Masterform in ihrer ursprünglichen Schablone ab.',
		'Select a standalone shape or copy its original master.':
			'Wählen Sie eine eigenständige Form oder kopieren Sie ihre ursprüngliche Masterform.',
		'Select a standalone 2D shape without a layer.':
			'Wählen Sie eine eigenständige 2D-Form ohne Ebene.',
		'The selected shape needs a positive width and height.':
			'Die ausgewählte Form benötigt eine positive Breite und Höhe.',
		'The shape and stencil use different Visio namespaces.':
			'Form und Schablone verwenden unterschiedliche Visio-Namespaces.',
		'The selected shape uses a non-default font.':
			'Die ausgewählte Form verwendet eine andere Schriftart als die Standardschrift.',
		'Another master entry uses this master part.':
			'Ein anderer Mastereintrag verwendet diesen Masterteil.',
	},
	es: {
		OK: 'Aceptar',
		Cancel: 'Cancelar',
		'Close {name}': 'Cerrar {name}',
		'Legacy binary stencils (.vss, .vst) are not supported.':
			'No se admiten las plantillas binarias antiguas (.vss, .vst).',
		'This file has no shapes.': 'Este archivo no tiene formas.',
		'This file is too large to open.': 'Este archivo es demasiado grande para abrirlo.',
		'Missing ZIP end record': 'Falta el registro de fin del ZIP',
		'Another part uses this master part.': 'Otra parte usa esta parte maestra.',
		'Edit Stencil': 'Editar plantilla de formas',
		'Turn on Edit Stencil to change its masters.':
			'Active la edición de la plantilla para cambiar sus formas maestras.',
		'Enter a name.': 'Escriba un nombre.',
		'The stencil exceeded the isolated editing time limit.':
			'La edición aislada de la plantilla superó el límite de tiempo.',
		'The Visio stencil worker failed.': 'Error en el proceso de la plantilla de Visio.',
		'New Stencil': 'Nueva plantilla de formas',
		'Stencil name': 'Nombre de la plantilla',
		'Untitled Stencil': 'Plantilla sin título',
		'Stencil actions': 'Acciones de la plantilla',
		'Save Stencil': 'Guardar plantilla de formas',
		'Add Selected Shape': 'Agregar forma seleccionada',
		'Master name': 'Nombre de la forma maestra',
		'New Master': 'Nueva forma maestra',
		'Rename Master': 'Cambiar nombre de forma maestra',
		'Delete Master': 'Eliminar forma maestra',
		'Edit Master': 'Editar forma maestra',
		'Default shape text': 'Texto predeterminado de la forma',
		'My Shapes': 'Mis formas',
		'Add to My Shapes': 'Agregar a Mis formas',
		'Remove from My Shapes': 'Quitar de Mis formas',
		'Select one editable 2D shape in a drawing.': 'Seleccione una forma 2D editable en un dibujo.',
		'Open the source drawing to copy this master.':
			'Abra el dibujo de origen para copiar esta forma maestra.',
		'Select a top-level shape with editable plain text.':
			'Seleccione una forma de nivel superior con texto sin formato editable.',
		'Enter a master name between 1 and 255 characters.':
			'Escriba un nombre de forma maestra de entre 1 y 255 caracteres.',
		'A master already has that name.': 'Ya existe una forma maestra con ese nombre.',
		'Enter plain text of up to 4096 characters.':
			'Escriba texto sin formato de hasta 4096 caracteres.',
		'Close another stencil before opening or editing this one.':
			'Cierre otra plantilla antes de abrir o editar esta.',
		'Editing requires a .vssx or .vssm stencil.':
			'La edición requiere una plantilla .vssx o .vssm.',
		'This master depends on another master in its original stencil.':
			'Esta forma maestra depende de otra forma maestra en su plantilla original.',
		'Select a standalone shape or copy its original master.':
			'Seleccione una forma independiente o copie su forma maestra original.',
		'Select a standalone 2D shape without a layer.':
			'Seleccione una forma 2D independiente sin capa.',
		'The selected shape needs a positive width and height.':
			'La forma seleccionada necesita un ancho y un alto positivos.',
		'The shape and stencil use different Visio namespaces.':
			'La forma y la plantilla usan distintos espacios de nombres de Visio.',
		'The selected shape uses a non-default font.':
			'La forma seleccionada usa una fuente distinta de la predeterminada.',
		'Another master entry uses this master part.': 'Otra entrada maestra usa esta parte maestra.',
	},
	fr: {
		OK: 'Valider',
		Cancel: 'Annuler',
		'Close {name}': 'Fermer {name}',
		'Legacy binary stencils (.vss, .vst) are not supported.':
			'Les anciens gabarits binaires (.vss, .vst) ne sont pas pris en charge.',
		'This file has no shapes.': 'Ce fichier ne contient aucune forme.',
		'This file is too large to open.': 'Ce fichier est trop volumineux pour être ouvert.',
		'Missing ZIP end record': 'L’enregistrement de fin du ZIP est manquant',
		'Another part uses this master part.':
			'Une autre partie utilise cette partie de forme de base.',
		'Edit Stencil': 'Modifier le gabarit',
		'Turn on Edit Stencil to change its masters.':
			'Activez la modification du gabarit pour changer ses formes de base.',
		'Enter a name.': 'Saisissez un nom.',
		'The stencil exceeded the isolated editing time limit.':
			'La modification isolée du gabarit a dépassé le délai autorisé.',
		'The Visio stencil worker failed.': 'Le traitement du gabarit Visio a échoué.',
		'New Stencil': 'Nouveau gabarit',
		'Stencil name': 'Nom du gabarit',
		'Untitled Stencil': 'Gabarit sans titre',
		'Stencil actions': 'Actions du gabarit',
		'Save Stencil': 'Enregistrer le gabarit',
		'Add Selected Shape': 'Ajouter la forme sélectionnée',
		'Master name': 'Nom de la forme de base',
		'New Master': 'Nouvelle forme de base',
		'Rename Master': 'Renommer la forme de base',
		'Delete Master': 'Supprimer la forme de base',
		'Edit Master': 'Modifier la forme de base',
		'Default shape text': 'Texte par défaut de la forme',
		'My Shapes': 'Mes formes',
		'Add to My Shapes': 'Ajouter à Mes formes',
		'Remove from My Shapes': 'Retirer de Mes formes',
		'Select one editable 2D shape in a drawing.':
			'Sélectionnez une forme 2D modifiable dans un dessin.',
		'Open the source drawing to copy this master.':
			'Ouvrez le dessin source pour copier cette forme de base.',
		'Select a top-level shape with editable plain text.':
			'Sélectionnez une forme de premier niveau avec du texte brut modifiable.',
		'Enter a master name between 1 and 255 characters.':
			'Saisissez un nom de forme de base de 1 à 255 caractères.',
		'A master already has that name.': 'Une forme de base porte déjà ce nom.',
		'Enter plain text of up to 4096 characters.':
			'Saisissez au maximum 4 096 caractères de texte brut.',
		'Close another stencil before opening or editing this one.':
			'Fermez un autre gabarit avant d’ouvrir ou de modifier celui-ci.',
		'Editing requires a .vssx or .vssm stencil.':
			'La modification nécessite un gabarit .vssx ou .vssm.',
		'This master depends on another master in its original stencil.':
			'Cette forme de base dépend d’une autre forme de base dans son gabarit d’origine.',
		'Select a standalone shape or copy its original master.':
			'Sélectionnez une forme autonome ou copiez sa forme de base d’origine.',
		'Select a standalone 2D shape without a layer.':
			'Sélectionnez une forme 2D autonome sans calque.',
		'The selected shape needs a positive width and height.':
			'La forme sélectionnée doit avoir une largeur et une hauteur positives.',
		'The shape and stencil use different Visio namespaces.':
			'La forme et le gabarit utilisent des espaces de noms Visio différents.',
		'The selected shape uses a non-default font.':
			'La forme sélectionnée utilise une police différente de la police par défaut.',
		'Another master entry uses this master part.':
			'Une autre entrée utilise cette partie de forme de base.',
	},
	'zh-CN': {
		OK: '确定',
		Cancel: '取消',
		'Close {name}': '关闭{name}',
		'Legacy binary stencils (.vss, .vst) are not supported.':
			'不支持旧版二进制模具（.vss、.vst）。',
		'This file has no shapes.': '此文件不包含任何形状。',
		'This file is too large to open.': '此文件过大，无法打开。',
		'Missing ZIP end record': '缺少 ZIP 结束记录',
		'Another part uses this master part.': '另一个部件正在使用此主控形状部件。',
		'Edit Stencil': '编辑模具',
		'Turn on Edit Stencil to change its masters.': '请开启模具编辑以更改主控形状。',
		'Enter a name.': '请输入名称。',
		'The stencil exceeded the isolated editing time limit.': '模具已超过隔离编辑时限。',
		'The Visio stencil worker failed.': 'Visio 模具后台处理失败。',
		'New Stencil': '新建模具',
		'Stencil name': '模具名称',
		'Untitled Stencil': '未命名模具',
		'Stencil actions': '模具操作',
		'Save Stencil': '保存模具',
		'Add Selected Shape': '添加所选形状',
		'Master name': '主控形状名称',
		'New Master': '新建主控形状',
		'Rename Master': '重命名主控形状',
		'Delete Master': '删除主控形状',
		'Edit Master': '编辑主控形状',
		'Default shape text': '默认形状文本',
		'My Shapes': '我的形状',
		'Add to My Shapes': '添加到我的形状',
		'Remove from My Shapes': '从我的形状中移除',
		'Select one editable 2D shape in a drawing.': '请在绘图中选择一个可编辑的二维形状。',
		'Open the source drawing to copy this master.': '请打开源绘图以复制此主控形状。',
		'Select a top-level shape with editable plain text.': '请选择包含可编辑纯文本的顶层形状。',
		'Enter a master name between 1 and 255 characters.': '请输入 1 到 255 个字符的主控形状名称。',
		'A master already has that name.': '已有同名主控形状。',
		'Enter plain text of up to 4096 characters.': '请输入不超过 4096 个字符的纯文本。',
		'Close another stencil before opening or editing this one.':
			'请先关闭另一个模具，然后打开或编辑此模具。',
		'Editing requires a .vssx or .vssm stencil.': '编辑需要 .vssx 或 .vssm 模具。',
		'This master depends on another master in its original stencil.':
			'此主控形状依赖原始模具中的另一个主控形状。',
		'Select a standalone shape or copy its original master.':
			'请选择一个独立形状，或复制其原始主控形状。',
		'Select a standalone 2D shape without a layer.': '请选择一个不在图层上的独立二维形状。',
		'The selected shape needs a positive width and height.': '所选形状的宽度和高度必须为正值。',
		'The shape and stencil use different Visio namespaces.':
			'形状和模具使用不同的 Visio 命名空间。',
		'The selected shape uses a non-default font.': '所选形状使用非默认字体。',
		'Another master entry uses this master part.': '另一个主控形状条目正在使用此部件。',
	},
} satisfies Record<VisioLocale, Record<string, string>>;
