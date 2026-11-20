# Migrating to 2.0

## Excel calculation (`ooxml-core/xlsx`)

### Each formula has its own work budget

A recalculation no longer shares one work budget among its formulas. Each formula may spend
about three seconds of priced work (matrix products, `LAMBDA` calls, repeated scans, sorting,
pattern matching); one past it returns `#NUM!`, usually before any of the work runs, and every
other formula still computes. A whole recalculation is bounded only when you ask for it, and
then it stops between formulas instead of turning later cells into errors.

Before:

```ts
const engine = createCalcEngine(workbook);
engine.recalculateAll(); // past one shared budget, every later formula was #NUM!
```

After:

```ts
const engine = createCalcEngine(workbook, { timeLimitMs: 10_000 });
engine.recalculateAll();
if (engine.stopped) {
	// 'time', 'work' or 'resources': the formulas it did not reach keep their previous values,
	// and the next recalculateAll() computes everything again.
}
```

### `CalcEngine` has a `stopped` member

Code that implements the `CalcEngine` interface itself, such as a test double, must add it.

Before:

```ts
const engine: CalcEngine = { recalculateAll() {}, recalculateFrom() {} /* ... */ };
```

After:

```ts
const engine: CalcEngine = {
	recalculateAll() {},
	recalculateFrom() {},
	/* ... */
	stopped: undefined,
};
```
