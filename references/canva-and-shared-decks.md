# Canva, Google Slides, and shared decks

## Importing a pptx

Build a normal `.pptx`, then import it. In Canva, use `--canva` so the builder leaves about 30px of extra gap on a 1920-wide slide (0.21 in). Canva renders the same font taller than PowerPoint, and tight gaps collide after import.

Google Slides imports `.pptx` as well. The same file is the source. After import, check type size and line breaks again. Do not assume the import preserved every box.

## Fonts

Use a font the target account actually has. Be Vietnam Pro is the default because it covers Vietnamese and is a free Google font. Anton is only for short Latin display text. Before you hand the file to Canva or Slides, confirm those faces exist in that account. If they do not, change the theme to a face that does, rebuild, and run the checker. A substituted font will reflow the lines and can drop diacritics.

The font files in `fonts/` are for local render and for installing on a machine. They are not a promise that every Canva team already has them.

## Fixing a deck you did not design

Keep the author's images, crops, and fonts unless they asked for a redesign. Change the type size, the line breaks, and the words. Do not swap the photography because you prefer another picture.

If the file is large, make a light copy with `--max-image-px 1300` and send that for review. Keep the full-size file if they still need it.

## Shared or live decks

Do not edit a live shared deck without a clear yes.

1. Check the latest edit time and the page count before you touch anything. If someone else edited it after you last looked, start from that version.
2. Do the work in a copy.
3. Show a before and after of the pages you want to change.
4. Change only the pages that were approved. Leave the rest of the file as you found it.
5. Never overwrite a teammate's file with your copy.
6. After the approved pages are in, check the page count again. An import that dropped or duplicated a slide is not done.

Transitions and Morph are off unless someone asked. A Morph only works when both slides contain the same picture object, not two separately placed crops.
