import { mountCloudFilesPlace } from 'ooxml-ui/pptx';
import type { CloudFileProvider, CloudFilesPlace, CloudOpenedFile } from 'ooxml-ui/pptx';
import React, { useEffect, useRef } from 'react';

import { useTranslation } from '../../i18n/use-translation';

/** One File > Open place for a cloud provider: the shared `pptx-ui-cloud-files` browser. */
export function CloudFilesPlaceView({
	provider,
	onOpen,
}: {
	provider: CloudFileProvider;
	onOpen: (file: CloudOpenedFile) => void;
}): React.ReactElement {
	const { t, locale } = useTranslation();
	const hostRef = useRef<HTMLDivElement>(null);
	const placeRef = useRef<CloudFilesPlace | null>(null);
	const options = { provider, translate: t, locale, onOpen };
	const latest = useRef(options);
	latest.current = options;
	useEffect(() => {
		const host = hostRef.current;
		if (!host) {
			return;
		}
		placeRef.current = mountCloudFilesPlace(host, latest.current);
		return () => {
			placeRef.current?.destroy();
			placeRef.current = null;
		};
	}, [provider]);
	useEffect(() => {
		placeRef.current?.update(latest.current);
	});
	return <div ref={hostRef} data-pptx-cloud-place={provider.id} className='mt-6 max-w-[900px]' />;
}
