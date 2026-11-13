import { elementRenderErrorContext, reportElementRenderError } from 'ooxml-ui/pptx';
import type {
	ElementRenderErrorContext as SharedElementRenderErrorContext,
	ElementRenderErrorHandler as SharedElementRenderErrorHandler,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React, { createContext, useContext, useMemo } from 'react';

/** Identifies the element omitted after a React client render or lifecycle error. */
export type ElementRenderErrorContext = SharedElementRenderErrorContext;

export type ElementRenderErrorHandler = SharedElementRenderErrorHandler;

interface ReportingScope {
	slideId?: string;
	onElementRenderError?: ElementRenderErrorHandler;
}

/** @internal Carries reporting through group children and reflection clones. */
export const ElementRenderErrorReportingContext = createContext<ReportingScope | undefined>(
	undefined,
);

interface ElementErrorBoundaryProps extends ReportingScope {
	element: PptxElement;
	children: React.ReactNode;
}

interface BoundaryState {
	element: PptxElement;
	failed: boolean;
}

class ElementErrorBoundaryState extends React.Component<ElementErrorBoundaryProps, BoundaryState> {
	state: BoundaryState = { element: this.props.element, failed: false };

	static getDerivedStateFromProps(
		props: ElementErrorBoundaryProps,
		state: BoundaryState,
	): Partial<BoundaryState> | null {
		return props.element !== state.element ? { element: props.element, failed: false } : null;
	}

	static getDerivedStateFromError(): Partial<BoundaryState> {
		return { failed: true };
	}

	componentDidCatch(error: Error): void {
		const { element, slideId, onElementRenderError } = this.props;
		// The shared payload and safe reporter: a throwing host reporter falls back to console.error.
		reportElementRenderError(
			onElementRenderError,
			error,
			elementRenderErrorContext(element, slideId),
		);
	}

	render(): React.ReactNode {
		return this.state.failed ? null : this.props.children;
	}
}

/** Gives every element boundary below a viewer-wide reporter (PowerPointViewer's prop). */
export function ElementRenderErrorScope({
	onElementRenderError,
	children,
}: {
	onElementRenderError?: ElementRenderErrorHandler | undefined;
	children: React.ReactNode;
}): React.ReactElement {
	const inherited = useContext(ElementRenderErrorReportingContext);
	const reporter = onElementRenderError ?? inherited?.onElementRenderError;
	const scope = useMemo(
		() => ({ slideId: inherited?.slideId, onElementRenderError: reporter }),
		[inherited?.slideId, reporter],
	);
	return (
		<ElementRenderErrorReportingContext.Provider value={scope}>
			{children}
		</ElementRenderErrorReportingContext.Provider>
	);
}

/** Omits only the failed element; a replacement element object retries rendering. */
export function ElementErrorBoundary({
	element,
	slideId,
	onElementRenderError,
	children,
}: ElementErrorBoundaryProps): React.ReactElement {
	const inherited = useContext(ElementRenderErrorReportingContext);
	const effectiveSlideId = slideId ?? inherited?.slideId;
	const reporter = onElementRenderError ?? inherited?.onElementRenderError;
	const scope = useMemo(
		() => ({ slideId: effectiveSlideId, onElementRenderError: reporter }),
		[effectiveSlideId, reporter],
	);
	return (
		<ElementRenderErrorReportingContext.Provider value={scope}>
			<ElementErrorBoundaryState
				key={element.id}
				element={element}
				slideId={effectiveSlideId}
				onElementRenderError={reporter}
			>
				{children}
			</ElementErrorBoundaryState>
		</ElementRenderErrorReportingContext.Provider>
	);
}
