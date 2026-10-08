/**
 * PowerPoint names for the format-neutral SmartArt model in `diagram/model`.
 * Each alias binds the neutral type's raw-XML slot to `XmlObject`, its run
 * style slot to `TextStyle` and its 3D slots to the pptx 3D models, so every
 * `PptxSmartArt*` type keeps the shape it had before the move.
 *
 * @module pptx-types/smart-art-model
 */

import type {
	DiagramAlgorithmParameter,
	DiagramChoose,
	DiagramColorApplicationMethod,
	DiagramColorListMetadata,
	DiagramColorTransform,
	DiagramColorTransformLabel,
	DiagramConstraint,
	DiagramConstraintOperator,
	DiagramConstraintPointType,
	DiagramConstraintRelationship,
	DiagramConstraintTarget,
	DiagramDefinitionCategory,
	DiagramDefinitionMetadata,
	DiagramDefinitionText,
	DiagramForEach,
	DiagramHueDirection,
	DiagramIteratorAttributes,
	DiagramLayoutAlgorithm,
	DiagramLayoutCategory,
	DiagramLayoutDefinition,
	DiagramLayoutNode,
	DiagramLayoutNodeShape,
	DiagramLayoutPreset,
	DiagramLocalizedText,
	DiagramNode,
	DiagramNodeCustomLayout,
	DiagramNodeStyle,
	DiagramNodeTextParagraph,
	DiagramNodeTextParagraphItem,
	DiagramNodeTextRun,
	DiagramNumericRule,
	DiagramPresLayoutVars,
	DiagramQuickStyle,
	DiagramQuickStyleLabel,
	DiagramResolvedStyleRef,
	DiagramRoleColorList,
	DiagramShapeAdjustment,
	DiagramWhen,
} from '../../../diagram/index';
import type { XmlObject } from './common';
import type { TextStyle } from './text';
import type { Pptx3DScene, Pptx3DShape, Text3DStyle } from './three-d';

// The aliases below resolve to the neutral names, so the declaration bundle
// emits those names and consumers whose inferred types reach them (for example
// `PptxSmartArtNodeStyle` resolving to `DiagramNodeStyle`) must be able to
// import them from `ooxml-core/pptx` too. Re-export every neutral model type
// and the part-model names the pptx SmartArt types reference.
export type * from '../../../diagram/model/index';
export type {
	DiagramColorScheme,
	DiagramConnection,
	DiagramLayoutType,
	DiagramNodeCustomLayout,
	DiagramStyleIntensity,
} from '../../../diagram/types';

// Constraints and rules (CT_Constraint, CT_NumericRule).
export type PptxSmartArtConstraintRelationship = DiagramConstraintRelationship;
export type PptxSmartArtConstraintOperator = DiagramConstraintOperator;
export type PptxSmartArtConstraintPointType = DiagramConstraintPointType;
export type PptxSmartArtConstraintTarget = DiagramConstraintTarget;
export type PptxSmartArtConstraint = DiagramConstraint<XmlObject>;
export type PptxSmartArtNumericRule = DiagramNumericRule<XmlObject>;

// Layout definition (CT_DiagramDefinition and its layout-node tree).
export type PptxSmartArtLocalizedText = DiagramLocalizedText;
export type PptxSmartArtLayoutCategory = DiagramLayoutCategory;
export type PptxSmartArtAlgorithmParameter = DiagramAlgorithmParameter;
export type PptxSmartArtLayoutAlgorithm = DiagramLayoutAlgorithm;
export type PptxSmartArtIteratorAttributes = DiagramIteratorAttributes;
export type PptxSmartArtForEach = DiagramForEach<XmlObject>;
export type PptxSmartArtWhen = DiagramWhen<XmlObject>;
export type PptxSmartArtChoose = DiagramChoose<XmlObject>;
export type PptxSmartArtShapeAdjustment = DiagramShapeAdjustment;
export type PptxSmartArtLayoutNodeShape = DiagramLayoutNodeShape;
export type PptxSmartArtLayoutNode = DiagramLayoutNode<XmlObject>;
export type PptxSmartArtLayoutDefinition = DiagramLayoutDefinition<XmlObject>;

// Colour-transform and quick-style definitions.
export type PptxSmartArtDefinitionText = DiagramDefinitionText;
export type PptxSmartArtDefinitionCategory = DiagramDefinitionCategory;
export type PptxSmartArtColorApplicationMethod = DiagramColorApplicationMethod;
export type PptxSmartArtHueDirection = DiagramHueDirection;
export type PptxSmartArtColorListMetadata = DiagramColorListMetadata;
export type PptxSmartArtResolvedStyleRef = DiagramResolvedStyleRef;
export type PptxSmartArtQuickStyleLabel = DiagramQuickStyleLabel<
	Pptx3DScene,
	Pptx3DShape,
	Text3DStyle
>;
export type PptxSmartArtColorStyleLabel = DiagramColorTransformLabel;
export type PptxSmartArtDefinitionMetadata = DiagramDefinitionMetadata;
export type PptxSmartArtColorTransform = DiagramColorTransform;
export type SmartArtRoleColorList = DiagramRoleColorList;
export type PptxSmartArtQuickStyle = DiagramQuickStyle<Pptx3DScene, Pptx3DShape, Text3DStyle>;

// Data-model nodes.
export type PptxSmartArtTextRun = DiagramNodeTextRun<TextStyle>;
export type PptxSmartArtTextParagraphItem = DiagramNodeTextParagraphItem<TextStyle>;
export type PptxSmartArtTextParagraph = DiagramNodeTextParagraph<TextStyle>;
export type PptxSmartArtNodeStyle = DiagramNodeStyle;
export type PptxSmartArtNode = DiagramNode<TextStyle>;
/**
 * Manual layout override for a `type="pres"` presentation point, read from its
 * `dgm:prSet` `cust*` attributes (see `DiagramNodeCustomLayout`).
 */
export type SmartArtNodeCustomLayout = DiagramNodeCustomLayout;

// Layout presets and variables.
export type SmartArtLayout = DiagramLayoutPreset;
export type PptxSmartArtPresLayoutVars = DiagramPresLayoutVars;
