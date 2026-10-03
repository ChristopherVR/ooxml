# Suite appearance sources

The shell remains one framework-neutral Visio viewer. No OOXML logic or framework
implementation was copied from another viewer.

Presentation rules were adapted from these read-only reference files:

- `ChristopherVR/pptx-viewer`, revision `7a27232ea0461c78df8c1ed9dd11636044d7c818`:
  `docs/.vitepress/theme/landing/LandingHero.vue`, `landing/landing.css`, and
  `docs/.vitepress/theme/custom.css`. Adaptations cover hero proportions, buttons,
  navigation, document columns and heading spacing. Visio copy retains beta limits.
- `ChristopherVR/docx-viewer`, revision `62bee302fae98b25c1860b2108bdd78f1627b267`:
  `demos/demo-vanilla/index.html` and `style.css`. Adaptations cover the local-file
  opening screen, titlebar size and Office surfaces. Visio has no new-document
  control because that capability is not implemented.

Both reference repositories use Apache-2.0. Existing root LICENSE and NOTICE apply
to the adapted presentation code.

The actual live reference sites use VitePress's Inter font. The Latin roman and
italic variable fonts were copied from the reference's installed VitePress font
assets, rather than declaring an unavailable display font. Inter is copyright
2016 The Inter Project Authors and is distributed under SIL Open Font License 1.1.
The exact license is retained at `assets/fonts/OFL.txt`.

`assets/viewer-preview.png` is a browser screenshot of this repository's original
sample in its actual viewer. It demonstrates the UI and does not establish native
Visio rendering parity.
