# OpenSlides

A standalone skill that teaches an AI agent how to build, fix, and review a slide deck. It works for PowerPoint (`.pptx`), Canva, and Google Slides. It is not tied to a brand or a project. Sample copy is a placeholder. Replace it.

The method is in [SKILL.md](SKILL.md). Agents should read that file first. [AGENTS.md](AGENTS.md) is the short checklist.

## Install

Copy this folder into the skills directory your agent reads:

- Cursor: `.cursor/skills/openslides/` or `~/.cursor/skills/openslides/`
- Claude Code: `~/.claude/skills/openslides/`
- Codex, Grok Bot, and other agents: put the folder where that agent loads skills, or point the agent at `SKILL.md`

The folder must keep `scripts/`, `templates/`, `references/`, and `fonts/` next to `SKILL.md`.

## Quickstart

```bash
python -m pip install -r scripts/requirements.txt
python scripts/build_deck.py templates/example-deck.yaml -o examples/example.pptx
python scripts/check_deck.py examples/example.pptx
bash scripts/render_preview.sh examples/example.pptx examples/preview
```

`examples/preview/` then holds `slide-01.png` onward and `contact-sheet.png`.

Other useful flags:

```bash
python scripts/build_deck.py deck.yaml -o deck.pptx --canva
python scripts/build_deck.py deck.yaml -o deck.pptx --theme templates/theme-ink.yaml
python scripts/build_deck.py deck.yaml -o deck-lite.pptx --max-image-px 1300
python scripts/check_deck.py deck.pptx --require-notes --check-links
```

Previews need LibreOffice and Poppler (`soffice` and `pdftoppm`):

```bash
sudo apt-get install -y libreoffice-impress poppler-utils
```

## What you get

- `scripts/build_deck.py` writes an editable 16:9 `.pptx` (13.333 × 7.5 in) from YAML or JSON. Each text element is its own text box.
- `scripts/check_deck.py` flags small type, crowded or long bullets, extra numbers, overflow, overlaps, low contrast, groups, bad Vietnamese fonts, and broken links.
- `scripts/render_preview.sh` writes a PNG per slide and a contact sheet.
- `templates/example-deck.yaml` is an 8-slide tour of the layouts. `theme-default.yaml` and `theme-ink.yaml` are the themes.
- `fonts/` contains Be Vietnam Pro and Anton, both under the SIL Open Font License (`fonts/OFL.txt`).

Layouts: `cover`, `section`, `bullets`, `image-left`, `image-right`, `big-number`, `comparison`, `quote`, `closing`, `sources`.

## Tests

```bash
python -m pytest -q
```

GitHub Actions runs the same command on push and pull request.

## License

MIT. See [LICENSE](LICENSE). The fonts keep their own license in `fonts/OFL.txt`.
