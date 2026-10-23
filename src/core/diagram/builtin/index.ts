// `ooxml-core/diagram/builtin`: the built-in SmartArt layout definitions and the default colour and
// quick-style definitions, as a separate entry because the layouts are about 350 KB gzipped. The
// main `diagram` entry never imports it; a product loads it when it inserts or re-lays-out a
// diagram.
export * from './layouts';
export * from './definitions';
