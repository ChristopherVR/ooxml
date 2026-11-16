// `ooxml-core/diagram/builtin`: the built-in SmartArt layout definitions, the default colour and
// quick-style definitions and the full built-in colour and quick-style catalogue, as a separate
// entry because the layouts are about 350 KB gzipped. The main `diagram` entry never imports it;
// a product loads it when it inserts, re-lays-out or restyles a diagram.
export * from './layouts';
export * from './definitions';
export * from './style-definitions';
