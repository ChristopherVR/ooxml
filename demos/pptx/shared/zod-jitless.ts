// Imported first by every demo entry. The published site's Content Security Policy never allows
// `eval`, and Zod probes for it with `Function('')` the first time it builds an object schema,
// which the browser reports as a violation. Turning Zod's JIT off before any schema exists skips
// the probe; validation works the same. Zod reads this object whether it loads before or after.
const zod = globalThis as { __zod_globalConfig?: { jitless?: boolean } };
(zod.__zod_globalConfig ??= {}).jitless = true;

export {};
