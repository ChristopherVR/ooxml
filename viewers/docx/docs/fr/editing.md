# Édition du texte

Chaque framework monte le même éditeur en composant web. Les commandes, les raccourcis clavier, la recherche, les métadonnées de langue et la conversion de documents n'ont qu'une seule implémentation.

## Rechercher et remplacer

Utilisez **Rechercher et remplacer** ou **Ctrl+F** (**Commande+F** sur macOS) lorsque l'éditeur a le focus. Saisissez une expression littérale, puis utilisez **Suivant** ou **Précédent** pour parcourir les correspondances. **Respecter la casse** contrôle la sensibilité à la casse. La touche Échap ferme le panneau et rend le focus au document.

La recherche couvre les séquences de texte mises en forme adjacentes au sein d'un paragraphe. Elle ne trouve pas de correspondance à cheval sur plusieurs paragraphes, cellules de tableau ou sauts de ligne forcés. Elle évite les correspondances qui scinderaient un graphème Unicode, par exemple une partie d'une séquence combinante ou d'un emoji composé. Les recherches sont littérales ; les expressions régulières et la correspondance linguistique propre à une langue ne sont pas prises en charge.

**Remplacer** modifie la correspondance courante ; **Tout remplacer** modifie toutes les correspondances du document en une seule opération annulable. Le texte de remplacement reprend la mise en forme du premier caractère trouvé. Le texte situé hors de la correspondance conserve sa mise en forme. La recherche fonctionne en lecture seule ; le remplacement y est désactivé. Les contrôles de préservation DOCX existants s'appliquent toujours à l'export de modifications effectuées autour d'un contenu source non pris en charge.

## Documents multilingues

La direction d'un paragraphe et la langue du texte sont des propriétés distinctes. Utilisez **Direction du paragraphe** pour une mise en page de gauche à droite ou de droite à gauche, et **Direction du texte** pour forcer explicitement une direction au sein d'un paragraphe. Les commandes de langue enregistrent les balises de langue de Word pour l'écriture standard, l'Asie orientale et les écritures complexes. Les métadonnées de langue ne traduisent pas le texte et n'installent pas de correcteur orthographique.

L'éditeur préserve le texte Unicode, y compris l'arabe, l'hébreu, le CJK, les caractères combinants et les emoji. Les polices du navigateur, la mise en forme des glyphes et la disposition bidirectionnelle déterminent l'apparence à l'écran. Ces commandes n'établissent ni une typographie des écritures complexes équivalente à celle de Word, ni sa pagination, ni le comportement des méthodes de saisie (IME) du système d'exploitation.

Le décompte des mots utilise la segmentation Unicode des mots du navigateur, y compris la segmentation par dictionnaire pour les écritures sans espaces lorsqu'elle est disponible. Les décomptes peuvent différer selon les versions d'ICU et par rapport à Microsoft Word. Les navigateurs plus anciens utilisent une solution de repli moins précise.

## Intégration de la co-édition

Le [guide de collaboration](/fr/collaboration) décrit le protocole d'édition versionné et la façon dont une application relie l'éditeur partagé à son autorité. Il s'agit d'une API d'intégration, pas d'un service de collaboration hébergé. Le contrôle d'accès, le réseau, le stockage et la politique de reconnexion relèvent de l'application.

## Langue de l'interface

Définissez la propriété `locale` de l'éditeur partagé (ou la prop `locale` de l'adaptateur) sur `en`, `fr`, `de`, `es` ou `zh-CN`. Les variantes régionales et d'écriture se résolvent vers la langue prise en charge la plus proche (`fr-CA`, `de-DE`, `es-MX`, `zh-Hans`, `zh-SG`) ; les locales non prises en charge, dont le chinois traditionnel (`zh-TW`, `zh-Hant`), reviennent à l'anglais. `EDITOR_LOCALES` liste les codes et `normalizeEditorLocale()` y ramène n'importe quelle balise. Chaque locale définit toutes les clés de chaînes (vérifié à la compilation et par des tests). Les chaînes allemandes, espagnoles et en chinois simplifié n'ont pas été relues par des locuteurs natifs. Changer la langue de l'interface préserve le document courant, la sélection, l'historique d'annulation et les métadonnées de langue du document.

## Styles de paragraphe importés

Le sélecteur de styles de l'onglet Accueil du ruban liste les styles de paragraphe du DOCX chargé. L'éditeur résout l'alignement, la direction, l'espacement et le retrait pris en charge à travers les valeurs par défaut du document et la chaîne d'héritage du style sélectionné. La mise en forme directe du paragraphe reste prioritaire. Choisir un style modifie sa référence, sans aplatir les propriétés héritées dans chaque paragraphe.

La partie de styles d'origine est préservée à l'enregistrement du document chargé. La création ou la modification de définitions de styles, les styles de caractère, la mise en forme héritée du texte et la résolution du thème ne sont pas encore pris en charge. L'export d'un nouveau paquet dont le catalogue de styles diffère des valeurs par défaut intégrées est rejeté ; utilisez Enregistrer sur le document chargé pour préserver ses définitions de styles d'origine.
