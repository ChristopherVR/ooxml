import {
	buildEquationBlockStyle,
	buildTextBlockStyle,
	convertOmmlToMathMl,
	EQUATION_FONT_FAMILY,
	isEquationOnlyText,
	sanitizeMathMl,
} from 'ooxml-ui/pptx';
import type { OmmlNode } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';
import React from 'react';

import { DEFAULT_TEXT_COLOR } from '../../constants';
import { cn } from '../../utils';

/**
 * A pure equation box (OMML and nothing else) laid out by the shared
 * {@link buildEquationBlockStyle} descriptor, the one Vue, Svelte and Vanilla
 * paint: the shape's own text body (insets, anchor, autofit) with each
 * equation row centred. Prose mixed with a formula is NOT this path: it keeps
 * the paragraph path, where the equation is a run in its authored position.
 *
 * Returns `null` when the element is not an equation-only body.
 */
export function renderEquationBox(
	el: PptxElement,
	interactive: boolean,
): React.ReactElement | null {
	if (!hasTextProperties(el) || !isEquationOnlyText(el.textSegments)) {
		return null;
	}
	// The rows are the same flex items the other bindings paint: no wrapper span
	// between the block and the maths, or its line box adds a strut of height.
	const equations = (el.textSegments ?? []).flatMap((segment, index) => {
		const mathml = segment.equationXml ? convertOmmlToMathMl(segment.equationXml as OmmlNode) : '';
		if (!mathml) {
			return [];
		}
		const math = (
			<span
				style={{
					display: 'inline-block',
					verticalAlign: 'middle',
					fontFamily: EQUATION_FONT_FAMILY,
				}}
				dangerouslySetInnerHTML={{ __html: sanitizeMathMl(mathml) }}
			/>
		);
		if (!segment.equationNumber) {
			return [<React.Fragment key={`${el.id}-eq-${index}`}>{math}</React.Fragment>];
		}
		const label = `(${segment.equationNumber})`;
		return [
			<span
				key={`${el.id}-eq-${index}`}
				style={{
					display: 'flex',
					justifyContent: 'space-between',
					alignItems: 'center',
					width: '100%',
				}}
			>
				<span style={{ visibility: 'hidden', whiteSpace: 'nowrap' }}>{label}</span>
				<span style={{ flex: 1, textAlign: 'center' }}>{math}</span>
				<span style={{ whiteSpace: 'nowrap', fontFamily: EQUATION_FONT_FAMILY }}>{label}</span>
			</span>,
		];
	});
	if (equations.length === 0) {
		return null;
	}
	const style = buildEquationBlockStyle(
		buildTextBlockStyle(el, {
			fallbackColor: DEFAULT_TEXT_COLOR,
			bodyLayout: true,
			pxLengths: true,
		}),
	) as React.CSSProperties;
	return (
		<div
			className={cn('relative z-10 w-full h-full', interactive ? '' : 'pointer-events-none')}
			style={style}
		>
			{equations}
		</div>
	);
}
