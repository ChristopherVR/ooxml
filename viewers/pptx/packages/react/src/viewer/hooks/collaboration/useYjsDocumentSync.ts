import type { CollaborationWriteBackState } from 'ooxml-core/collab';
import type {
	CollabLoadOrigin,
	CollaborationConfig,
	CollaborationLivePatcher,
	YTransactionLike,
} from 'ooxml-ui/pptx';
import {
	configureCollaborationMedia,
	createRoomJoinTracker,
	reconcileSlidesInYDoc,
	LOCAL_SYNC_ORIGIN,
	readSlidesFromYDoc,
	observeYDocSlides,
	shouldRoomSlidesReplaceLoad,
} from 'ooxml-ui/pptx';
/** Synchronize viewer slides with the shared granular Yjs slide schema. */
import type { PptxElement, PptxHandlerSaveOptions, PptxSlide } from 'pptx-viewer-core';
import { useEffect, useRef, useState } from 'react';
import type { Doc as YDoc } from 'yjs';

import { useCollaborationWriteBack } from './useCollaborationWriteBack';
import { useExternalDocumentSync } from './useExternalDocumentSync';

export interface UseYjsDocumentSyncInput {
	doc: YDoc | null;
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	isConnected: boolean;
	isSynced?: boolean;
	config?: Pick<
		CollaborationConfig,
		| 'role'
		| 'onWriteBack'
		| 'writeBackDebounceMs'
		| 'writeBackRetry'
		| 'externalSession'
		| 'sessionIntent'
	>;
	getSourceBytes?: () => Uint8Array | null;
	getSaveOptions?: () => PptxHandlerSaveOptions;
	loadVersion?: number;
	loadOrigin?: CollabLoadOrigin;
	/** External readiness owns this channel together with full-deck writes. */
	livePatcher?: CollaborationLivePatcher;
	onReadOnlyChange?: (readOnly: boolean) => void;
	/** Receives the owner's write-back status (undefined when the session ends). */
	onWriteBackStatus?: (state: CollaborationWriteBackState | undefined) => void;
	/** The room's slides replaced the open deck (a join): the binding counts it as a presentation load. */
	onRoomReplacedDeck?: () => void;
	/** The URLs this viewer plays its audio and video from, so the room can carry their bytes. */
	mediaDataUrls?: ReadonlyMap<string, string>;
}

export function useYjsDocumentSync(input: UseYjsDocumentSyncInput): void {
	const {
		doc,
		slides,
		setSlides,
		isConnected,
		isSynced = true,
		config,
		loadVersion = 0,
		loadOrigin = 'user',
	} = input;
	const external = Boolean(config?.externalSession);
	const latestMedia = useRef(input);
	latestMedia.current = input;
	// Once per room, before any effect below reads or writes its slides.
	useEffect(() => {
		if (doc) {
			configureCollaborationMedia(doc, {
				localMediaUrl: (path) => latestMedia.current.mediaDataUrls?.get(path),
			});
		}
	}, [doc]);
	const scheduleWriteBack = useCollaborationWriteBack(input);
	useExternalDocumentSync(input, scheduleWriteBack);
	const lastSynced = useRef('');
	const superseded = useRef<PptxSlide[] | null>(null);
	const latestSlides = useRef(slides);
	latestSlides.current = slides;
	const lastLoad = useRef(loadVersion);
	const revision = useRef(0);
	const [join] = useState(createRoomJoinTracker);
	const latestInput = useRef(input);
	latestInput.current = input;
	const replacedDeck = (): void => latestInput.current.onRoomReplacedDeck?.();

	useEffect(() => {
		if (external || !doc) {
			return;
		}
		lastSynced.current = '';
		superseded.current = null;
		revision.current++;
		join.attached(latestInput.current.config?.sessionIntent);
	}, [doc, external, join]);

	useEffect(() => {
		if (external || !doc || !isConnected) {
			return;
		}
		const adopt = (): void => {
			const remote = readSlidesFromYDoc(doc);
			if (!remote.length) {
				return;
			}
			const serialized = JSON.stringify(remote);
			if (serialized === lastSynced.current) {
				return;
			}
			lastSynced.current = serialized;
			superseded.current = latestSlides.current;
			revision.current++;
			setSlides(remote);
			if (join.adopted()) {
				replacedDeck();
			}
		};
		const unobserve = observeYDocSlides(doc, (_events, transaction?: YTransactionLike) => {
			if (transaction?.origin === LOCAL_SYNC_ORIGIN) {
				return;
			}
			adopt();
			scheduleWriteBack();
		});
		adopt();
		const revokeWrites = (): void => {
			revision.current++;
		};
		return () => {
			unobserve();
			revokeWrites();
		};
	}, [external, doc, isConnected, setSlides, scheduleWriteBack, join]);

	useEffect(() => {
		if (external || loadVersion === lastLoad.current) {
			return;
		}
		lastLoad.current = loadVersion;
		join.loaded(loadOrigin);
		if (!doc || !isConnected) {
			return;
		}
		const room = readSlidesFromYDoc(doc);
		if (shouldRoomSlidesReplaceLoad(loadOrigin, room.length)) {
			lastSynced.current = JSON.stringify(room);
			superseded.current = latestSlides.current;
			revision.current++;
			setSlides(room);
			join.adopted();
			replacedDeck();
		}
	}, [external, loadVersion, loadOrigin, doc, isConnected, setSlides, join]);

	useEffect(() => {
		if (
			external ||
			!doc ||
			!isConnected ||
			!isSynced ||
			config?.role === 'viewer' ||
			!slides.length ||
			slides === superseded.current ||
			!join.mayWrite()
		) {
			return;
		}
		const serialized = JSON.stringify(slides);
		if (serialized === lastSynced.current) {
			return;
		}
		let cancelled = false;
		const token = revision.current;
		void (async () => {
			const Y = await import('yjs');
			if (cancelled || token !== revision.current || loadVersion !== lastLoad.current) {
				return;
			}
			reconcileSlidesInYDoc(slides, doc, {
				createMap: () => new Y.Map(),
				createArray: () => new Y.Array(),
				createText: () => new Y.Text(),
			});
			lastSynced.current = serialized;
			join.wrote();
			scheduleWriteBack();
		})();
		return () => {
			cancelled = true;
		};
	}, [
		external,
		doc,
		slides,
		isConnected,
		isSynced,
		config?.role,
		loadVersion,
		scheduleWriteBack,
		join,
	]);
}
