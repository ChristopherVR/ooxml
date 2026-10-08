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
	for?: DiagramConstraintRelationship | undefined;
	forName?: string | undefined;
	pointType?: DiagramConstraintPointType | undefined;
}

/** Editable DiagramML CT_Constraint. */
export interface DiagramConstraint<R = unknown> extends DiagramConstraintTarget {
	type: string;
	referenceType?: string | undefined;
	referenceFor?: DiagramConstraintRelationship | undefined;
	referenceForName?: string | undefined;
	referencePointType?: DiagramConstraintPointType | undefined;
	operator?: DiagramConstraintOperator | undefined;
	value?: number | undefined;
	factor?: number | undefined;
	/** Original constraint retained for foreign attributes and extension content. */
	rawXml?: R | undefined;
}

/** Editable DiagramML CT_NumericRule. */
export interface DiagramNumericRule<R = unknown> extends DiagramConstraintTarget {
	type: string;
	value?: number | undefined;
	factor?: number | undefined;
	max?: number | undefined;
	/** Original rule retained for foreign attributes and extension content. */
	rawXml?: R | undefined;
}
