# Attributions

Font Manager for PowerPoint is built on other people's work. This file lists what that work is, who did
it, and what it is doing here.

It is generated — the master lists live in the `stoatworks-backend` repo and are
pushed out by `scripts/sync-attributions.py`. Edit it there, not here.

## Code we derived from other people's work

Someone else solved this first, and this project would not exist in its current form without their work.

### Font scanner and resolver — Stoatworks pptx-font-manager

<https://github.com/stoatworks-labs/pptx-font-manager>  
Licence: MIT  
Copyright: Stoatworks Labs

Same fleet, copied rather than shared: src/core, src/lib, src/platform and src/data are vendored from pptx-font-manager under vendor/pptx-font-manager via git subtree and imported unchanged through the @fm alias, so the add-in and the web tool cannot disagree about a deck. That brings the scanner, the resolver, the canvas width-probe font check, the metric-compatible substitutes and the Google Fonts, Fontsource and Adobe Fonts catalogues. The ribbon icons are rendered from pptx-font-manager's icon.svg.

## Third-party code this project uses

Libraries, SDKs and frameworks the project is built on or bundles.

### The npm ecosystem

<https://www.npmjs.com>  
Licence: predominantly MIT  
Copyright: the individual package authors

npm dependencies, resolved and pinned in the lockfile.

Build tooling, test runners and the libraries the front ends are assembled from. The exact set and versions for any build are in that repo's lockfile, which is the authoritative list.

The full transitive dependency set for any build is pinned in this repo's lockfile,
which is the authoritative list. What is named above is the layers a reader would
want to know about, not every package that has ever been resolved.

## Getting this wrong

If your work is here and the description is inaccurate, the licence is wrong, or you would rather not be listed — open an issue and it will be fixed.
