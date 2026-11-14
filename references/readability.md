# Readability

Look at the slide the way a person in the back row will see it. If you have to lean in, it is too small.

## Size

Build on a 13.333 × 7.5 in slide. That is the 16:9 frame people export as 1920 × 1080.

Body text defaults to **28pt**. At 96dpi that is about 37px, which is the size the writing rules call 36–40px. The floor is **22pt** (about 30px at 96dpi). The checker fails text below 22pt.

On a 1920-wide export (144 dpi) 28pt is 56 screen pixels. Use the point size on the 13.333 in slide. Do not shrink the type to chase a pixel number from a different density.

Titles are 36pt or larger and never smaller than the body. A big number is display size. Captions stay at 22pt or above. There is no small grey caption style.

## Weight and contrast

Use Medium or SemiBold for body copy. Regular is thin at projector distance. The default theme uses Be Vietnam Pro SemiBold for body and Be Vietnam Pro Bold for titles.

Put dark text on a light solid fill, or light text on a dark solid fill. The checker uses WCAG contrast: 4.5:1 under 24pt, 3:1 at 24pt and above. Theme colors in this repo clear 4.5:1 on the surface they are meant for. Do not recolor a word to grey to make it "secondary."

## Density

Three or four points. Then a new slide. Generous space between text boxes. The builder's `--canva` flag adds about 30px (0.21 in) because Canva draws these fonts taller than PowerPoint. See `canva-and-shared-decks.md`.

Do not put type on a photograph unless a solid panel covers the area behind the words. The checker flags text whose center sits on a picture with no solid shape in between.

## Diacritics

Set the font on the Latin, East Asian, and complex-script slots so a Vietnamese syllable does not fall back to a face that drops marks. The builder does this, and it sets `vi-VN` on runs that contain Vietnamese letters.

Anton, Impact, Bebas, and Arial Black do not cover Vietnamese. The checker fails that combination. Render the real line and read it. "Mỗi ý là một câu đủ nghĩa và giữ đủ dấu tiếng Việt." is the sample in the example deck.

## What to look for on the PNG

Overflow, boxes that collide, type the same color as its panel, a duplicated line, a broken mark, a typo, a widow word on its own line, and a title smaller than the body. Fix the cause. Do not nudge one box and leave the same mistake on the next slide.
