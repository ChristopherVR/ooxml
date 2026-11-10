// `pptx-react-viewer/no-i18next`: the API of `pptx-react-viewer`, for hosts that translate only
// through `PptxI18nProvider`. Nothing it imports reaches react-i18next or i18next, so a host can
// bundle it without either. Outside a provider the viewer shows its built-in English text and
// leaves File > Options > Language out.
export * from './public-api';
