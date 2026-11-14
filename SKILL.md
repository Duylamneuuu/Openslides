---
name: openslides
description: Use when building, fixing, or reviewing any slide deck (.pptx, Canva, or Google Slides).
---

# OpenSlides

Build, fix, or review a slide deck that a room can read. The file may be PowerPoint, Canva, or Google Slides. This skill has no brand kit. Examples are placeholders. Replace them.

Longer notes live in `references/writing-guide.md`, `references/readability.md`, `references/canva-and-shared-decks.md`, and `references/style-directions.md`.

## Before designing

1. Collect the facts first, each with a real source. Never invent a figure, a quote, or a citation. If a number has no source, leave it out.
2. Write the storyline. Each slide has one job the listener can repeat.
3. Gather real images before you lay anything out. A generated image is only for an object or a background, only with the user's OK, and never for a person.
4. If the look is not decided, offer 2 or 3 distinct directions on 2 or 3 sample slides. Wait for a choice before building the rest.
5. Make a page plan. For every slide, write the job, the layout, the image, and the source.

Use 5 to 7 layouts in a deck. Do not use one layout more than about 3 times. Do not put more than two similar slides in a row. Put a strong visual about every four slides.

## Readability

Judge the whole slide from the back of the room.

- The slide is 13.333 × 7.5 in (16:9, the usual 1920 × 1080 frame).
- Body text defaults to 28pt. That is about 37px in 96dpi CSS pixels. Do not set body text below 22pt (about 30px).
- Use Medium or SemiBold. Keep contrast high. Do not use small grey captions.
- Three or four points. Split a long idea onto another slide. Do not shrink the type, and do not cut words to make it fit.
- The title names the topic. It is never smaller than the body.
- Leave generous space. Do not put text on a busy photo unless a solid panel sits behind the words.
- For Canva, add about 30px of extra gap (`--canva`). Canva draws the same font taller.
- Render the deck and read it. If the deck uses Vietnamese, read the marks. A missing mark means the wrong font.

## Writing

Write plain, everyday, full sentences. No flourish, no slogans, no semicolon chains, no repeated "X, but Y", no decorative quotation marks.

A bullet is a bold label, a colon, and one plain sentence of about 10 to 20 words:

**One job:** Say the single thing this slide is for.

An analysis slide has 0 to 2 numbers. A big number belongs on a data slide, with a label and a caption, and it comes from a source. Keep each fact the same on every slide: same name, same figure, same spelling.

The slide holds the idea. Timing, asides, and extra detail go in the speaker notes. End with a sources slide of hyperlinks when the deck cites anything.

## Build

Everything stays editable. One text box per text element. No groups. Do not flatten text into a picture.

Use a font the target tool actually has. The default is Be Vietnam Pro, which covers Vietnamese, plus Anton for a short Latin display string such as a big number. If a run contains Vietnamese and the chosen font has no diacritics, the builder switches that run to Be Vietnam Pro.

When you fix someone else's deck, keep their images, crops, and fonts unless they asked you to redesign.

If the file is heavy, make a light copy with `--max-image-px 1300` so images are about 1300px on the long edge.

Add a transition or Morph only when asked. Morph needs the same picture object on both slides, not two crops.

## Check before sending

Render every slide and look. Fix overflow, overlaps, text that matches its background, duplicate lines, broken diacritics, typos, and orphan words. Check every number against its source. Check the links.

Deliver the preview images and the editable file. When you edited an existing deck, include a before and after.

## Shared and live decks

Do not edit a live shared deck without approval. Fix a copy, show the comparison, and change only the pages that were approved. Check the latest edit time and the page count first. Never overwrite a teammate. Check the page count again after you update.

## Feedback

Fix the cause on every slide, not only the one that was marked. When the note is a preference (tone, density, color), apply it to the rest of the deck and keep it for the next one.

## Layouts

| Layout | Use |
| --- | --- |
| `cover` | Title, optional kicker and subtitle |
| `section` | Section break |
| `bullets` | Title and up to four labeled sentences |
| `image-left`, `image-right` | Photograph and text |
| `big-number` | One sourced figure, a label, and a caption |
| `comparison` | Two columns |
| `quote` | A real quotation and its source, with no decorative quote marks |
| `closing` | The last point and the next step |
| `sources` | Names that are hyperlinks |

`templates/example-deck.yaml` is a filled spec. `templates/theme-default.yaml` and `templates/theme-ink.yaml` are the two themes.

## Workflow

1. Collect facts with sources, write one job per slide, gather real images, and make a page plan. If the look is open, propose 2 or 3 directions on sample slides first (`references/style-directions.md`, `references/writing-guide.md`).
2. Copy `templates/example-deck.yaml` and replace the placeholder copy. Keep `templates/theme-default.yaml`, or pass `templates/theme-ink.yaml`.
3. Build the editable file:

```bash
python scripts/build_deck.py deck.yaml -o deck.pptx
```

Add `--canva` when the file will be imported into Canva. Add `--theme templates/theme-ink.yaml` to swap the theme. Add `--max-image-px 1300` for a light copy.

4. Lint it. Do not deliver a file that fails:

```bash
python scripts/check_deck.py deck.pptx
```

Add `--require-notes` when every slide should have speaker notes. Add `--check-links` when the deck has URLs.

5. Render a PNG of every slide and a contact sheet, then look at each picture:

```bash
bash scripts/render_preview.sh deck.pptx preview/
```

6. Fix the root cause and run steps 4 and 5 again. Deliver the `.pptx` and the PNGs. When you edited an existing deck, deliver a before and after as well.

Fonts in `fonts/` are SIL Open Font License. `scripts/render_preview.sh` installs them for the current user before it renders. Install them yourself too if a preview substitutes a different face.
