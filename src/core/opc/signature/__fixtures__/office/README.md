# Office-signed packages

Packages that Microsoft Office signed, used to check that the platform-neutral signature verifier
accepts real Office signatures (`../../verify/office-signatures.test.ts`). They are test data from
Apache POI (`https://github.com/apache/poi`, `test-data/xmldsign/`, Apache License 2.0), copied
unchanged; the Git blob of each upstream file is listed so a copy can be checked against it.

| File                                             | Signed by                                  | Upstream blob                              |
| ------------------------------------------------ | ------------------------------------------ | ------------------------------------------ |
| `PPT2016withComment.pptx`                        | PowerPoint 2016 (Office 16.0), `CN=Test`   | `92154f9f0d69763e9402b4d6d67ffbd0ceed4282` |
| `hello-world-office-2010-technical-preview.docx` | Word 2010 (Office 14.0), `CN=Test`         | `cbd4277564bbf14009e01a1aa402672748dbf84b` |
| `office2007prettyPrintedRels.docx`               | Word 2007 (relationships pretty-printed)   | `ce8b940944a3b8c94976233e903ecc14423abfbb` |
| `hello-world-signed.xlsx`                        | Excel 2007 (Office 12.0)                   | `0d45c53ede3435b2062d12973436d8eef8c52e63` |

No signer here chains to a trust root this repository configures, so each verifies as
`verified-untrusted`. Visio signatures are produced by the tests themselves with a generated test
certificate (`src/core/opc/signature/test-support/sign.ts`), since no Visio-signed sample is
available under a compatible licence.
