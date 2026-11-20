# Generated signed packages

Packages signed by the test signer of `../../test-support/sign.ts` (`signForTest`), committed so
tests and browser specs that cannot run the node signing code (the `ooxml-ui` DOM tests and the
Playwright specs) still open a signed document. They hold only the public test certificate; the
private key is generated per run and never stored.

| File                 | Source                                                 | State                |
| -------------------- | ------------------------------------------------------ | -------------------- |
| `signed-drawing.vsdx` | `src/core/visio/__fixtures__/native-alt-text.vsdx`    | `verified-untrusted` |

Changing a signed part (invalid) or the signature method (unsupported) is done by the tests.
