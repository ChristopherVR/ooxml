// `pptx-react-viewer`. Outside any `PptxI18nProvider` the viewer translates through the host's
// react-i18next instance, which this import registers. `pptx-react-viewer/no-i18next` exports
// the same API without it.
import './viewer/i18n/react-i18next-source';

export * from './public-api';
