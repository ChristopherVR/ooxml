import { galleryTextLabel } from 'ooxml-ui/pptx';

import { useTranslation } from '../../i18n/use-translation';

export function useTranslateOr(): (key: string | undefined, fallback: string) => string {
	const { t } = useTranslation();
	return (key, fallback) => galleryTextLabel(t, key, fallback);
}
