# Art direction and the page plan

## When the look is not chosen

Do not build the whole deck. Offer 2 or 3 directions that differ in type, color, and density. Three shades of the same blue are one direction.

Show each direction on the same 2 or 3 slides: the cover, one text slide, and one visual slide (a photograph, a big number, or a comparison). Use real copy from the page plan, not lorem ipsum.

Name the directions in plain words. "Warm paper" is the default theme. "Ink" is `templates/theme-ink.yaml`. A third direction should change something a person can see, such as a poster scale for the titles or a quiet one-column pace. Say what you would keep and what you would drop.

Wait for a choice. Then build the rest of the deck in that direction only.

## Page plan

Write this before the full spec. One row per slide:

- Job: the single thing the listener should take away.
- Layout: one of the named layouts.
- Image: the real photograph, or none. Note if a generated object or background still needs a yes.
- Source: where the facts on that slide come from.

Then check the plan:

- 5 to 7 layouts in the deck.
- No layout more than about 3 times.
- No more than two similar slides in a row.
- A strong visual about every four slides.
- Every number has a source. Slides with no source have no number.

`templates/example-deck.yaml` is a layout tour, so it changes layout every slide. A real deck may repeat a layout. It still follows the rules above.

## Using a style skill with this one

If the user has a style skill or a brand skill, take palette, type mood, and image treatment from that skill. Keep OpenSlides for the story, the sentences, the type size, the editable file, and the checks.

Do not copy a brand out of an old deck or out of this repository. This repository does not have one. Put the chosen colors and font names in a theme file next to `templates/theme-default.yaml`, and keep body text at 28pt or above with contrast the checker will pass.

When the style skill and this skill disagree about type size or contrast, this skill wins. A beautiful slide that the back row cannot read is a failed slide.
