// Neutral DiagramML constraint and rule model (CT_Constraint, CT_NumericRule). `R` is the raw XML
// node a reader keeps for round-trip (pptx: its `XmlObject` tree); the layout engine never reads it.
// Moved from `pptx/core/types/smart-art-constraint-rules.ts`.

export type DiagramConstraintRelationship = 'self' | 'ch' | 'des';
export type DiagramConstraintOperator = 'none' | 'equ' | 'gte' | 'lte';
export type DiagramConstraintPointType =
	| 'all'
	| 'doc'
	| 'node'
	| 'norm'
	| 'nonNorm'
	| 'asst'
	| 'nonAsst'
	| 'parTrans'
	| 'pres'
	| 'sibTrans';

export interface DiagramConstraintTarget {
	for?: DiagramConstraintRelationship;
	forName?: string;
	pointType?: DiagramConstraintPointType;
}

/** Editable DiagramML CT_Constraint. */
export interface DiagramConstraint<R = unknown> extends DiagramConstraintTarget {
	type: string;
	referenceType?: string;
	referenceFor?: DiagramConstraintRelationship;
	referenceForName?: string;
	referencePointType?: DiagramConstraintPointType;
	operator?: DiagramConstraintOperator;
	value?: number;
	factor?: number;
	/** Original constraint retained for foreign attributes and extension content. */
	rawXml?: R;
}

/** Editable DiagramML CT_NumericRule. */
export interface DiagramNumericRule<R = unknown> extends DiagramConstraintTarget {
	type: string;
	value?: number;
	factor?: number;
	max?: number;
	/** Original rule retained for foreign attributes and extension content. */
	rawXml?: R;
}
