# Signature state decks

One deck per signature state a browser can reach, for `e2e/pptx/signature-state.spec.ts`. Each
file is named after the state the browser verifier (`inspectSignaturesFromBuffer`) gives it.

| File                      | How it is made                                                                                                           |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| `unsigned.pptx`           | A one-slide deck saved by the core, with no signature parts.                                                             |
| `verified-untrusted.pptx` | That deck signed with a self-signed certificate (`CN=OOXML Signature Fixture`). It verifies; a browser trusts no signer. |
| `invalid.pptx`            | The signed deck with a comment added to `ppt/slides/slide1.xml` after signing.                                           |
| `unsupported.pptx`        | The signed deck with its `SignatureMethod` changed to ECDSA-SHA256, which the verifier does not implement.               |

`viewers/pptx/scripts/make-signature-state-decks.mjs` writes all four and fails, writing nothing,
unless each reaches its state. The certificate and its private key are created in memory for one
run and never stored, so every run signs with a new key and changes the three signed decks:
regenerate only when they need to change, and commit the four together.

```bash
cd viewers/pptx && bun run scripts/make-signature-state-decks.mjs
```

The decks are not part of `generate-all.ts` (`fixtures:check` would see them change on every run)
and not in the core's `src/core/pptx/__tests__/fixtures/e2e` snapshot.
