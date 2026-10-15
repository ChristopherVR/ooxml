/**
 * review-comments-navigation.service.ts: the viewer's active slide as the Review tab's
 * Previous / Next / Delete comment commands need it. The active slide is a signal on the
 * viewer component itself, so the component registers it here (`attach`) and the ribbon's
 * Review section reads it without threading three more inputs through the ribbon.
 */
import { Injectable, signal } from '@angular/core';

@Injectable()
export class ReviewCommentsNavigationService {
	private goTo: (slideIndex: number) => void = () => {};

	/** The active slide index, kept in step by the viewer. */
	readonly currentSlide = signal(0);

	/** Called once by the viewer: how to move to a slide. */
	attach(goToSlide: (slideIndex: number) => void): void {
		this.goTo = goToSlide;
	}

	goToSlide(slideIndex: number): void {
		this.goTo(slideIndex);
	}
}
