# Security Policy

## Supported versions

`@christophervr/ooxml-core` is a single package. Security fixes target the **latest release**; there are no long-lived security branches. Upgrade to the latest version before reporting a suspected vulnerability, in case it has been fixed.

The package parses untrusted Office files (OOXML packages, XML parts, embedded media and metafiles, encrypted packages) and can write them back. Reports about how it handles hostile input are especially welcome: XML entity or DTD handling, zip bombs and path traversal in package parts, unbounded recursion or memory growth, unsafe hyperlink or URL schemes, and weaknesses in the encryption or signature code paths.

## Reporting a vulnerability

Please do **not** open a public issue. Report it through a [private GitHub security advisory](https://github.com/ChristopherVR/ooxml/security/advisories/new) and include:

- the affected area (`xml`, `opc`, `docx`, `pptx`, ...) and version,
- a minimal file or input that reproduces it, and the observed impact,
- whether you intend to disclose publicly, and when.

You will get an acknowledgement as soon as the maintainer can respond, and updates as the issue is investigated and fixed. Fixes ship in a new release with a note in the release description; reporters are credited unless they ask not to be.

## Scope

In scope: the code in this repository and the published package. Out of scope: vulnerabilities in the viewer applications' UI code (report those in `docx-viewer` or `pptx-viewer`), and in third-party dependencies (report those upstream; we will update once a fix exists).
