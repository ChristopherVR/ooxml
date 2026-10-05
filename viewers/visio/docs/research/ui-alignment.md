# PPTX visual alignment

The Visio workspace and documentation use the same warm-paper / dark-presenter
palette as the current `ChristopherVR/pptx-viewer` reference. This is a UI design
alignment, not a claim of equivalent editing or file-format capabilities.

## Primary references inspected

- `pptx-viewer/packages/shared/src/theme/presets.ts`: vermilion light/dark tokens
- `pptx-viewer/packages/shared/src/render/editor-chrome/`: shared compact chrome
- `pptx-viewer/DESIGN.md`: title/ribbon/rail/inspector/status dimensions
- `pptx-viewer/docs/.vitepress/theme/landing/landing.css` and `LandingHero.vue`:
  warm dot-grid canvas, rust display emphasis, monospaced labels and framed stage
- `pptx-viewer/docs/.vitepress/theme/custom.css`: coordinated documentation theme
- Live public PPTX landing and demo, inspected on 2026-10-02. The live demo
  reported build `771c48b`. Token source was also checked through GitHub's API.

Source repository: https://github.com/ChristopherVR/pptx-viewer
Reference site: https://christophervr.github.io/pptx-viewer/

## Deliberate scope

The app replaces its marketing sidebar with a document workspace: compact file
commands, actual page navigation, Home/View controls, canvas, inspector and zoom
status. Search, plain-text editing, history, layer visibility and exports continue
to use the shared controller and custom element. No fake drawing, collaboration,
presentation, print-dialog or formatting controls are added.

The site preserves the public-beta and placeholder-package status, local source
setup and honest rendering limitations. All assets are local. Appearance choices
are shared between the site and workspace through optional local storage. Desktop
chrome is compact; mobile controls retain 44px touch targets.

## Visual evidence

`tests/design-evidence.spec.ts` captures desktop/mobile and light/dark images of
the workspace, landing page, loaded embedded demo and guide using synthetic sample
content only. The verification workflow retains these images on success as well
as failure. Passing DOM/source checks alone is not visual verification: review
these Chromium screenshots for clipping, readable text, pane layout and spacing.
