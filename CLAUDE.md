# Shift Projects slide builder (instructions for Claude)

This repo is the `frontend-slides` skill, set up for **Shift Projects** decks. It is installed at `~/.claude/skills/frontend-slides`; users start it with `/frontend-slides`.

## What to do for a Shift Projects deck
1. Read `shift-projects-brand.md` in full, then clone `shift-projects-slide-library.html` (19 layouts). Never regenerate its CSS; edit content only. Skip style discovery.
2. Follow the voice rules in the brand file: clear, direct, people-first, no jargon, sentence case.
3. Render previews and PDFs at **1920x1080** (`bash scripts/export-pdf.sh deck.html`).
4. Save decks to the user's Downloads folder unless told otherwise.

## If they want it in Google Slides
Use the native template, not the HTML (HTML does not convert cleanly):
1. `import sys; sys.path.insert(0, '<skill>/scripts'); from shift_pptx import *`
2. `prs = open_template(); keep(prs, [1, 4, 8, 19])` picks and orders layouts by number (see the list in `shift-projects-brand.md`); repeat a number to duplicate that slide.
3. `set_text(prs.slides[i], 'Heading', '...')` edits text by shape name; `texts(slide)` lists the names.
4. Save to Downloads, then tell the user: drag the `.pptx` into Google Drive, right-click, Open with Google Slides. Details: `docs/google-slides.md`.
Do not add or invent brand elements. Photos are placeholders: tell the user to insert images over the grey tiles.

## First-time setup (a fresh machine)
Run `bash ~/.claude/skills/frontend-slides/setup.sh`. It checks for Node.js and Python, installs `python-pptx`, and downloads the PDF-export browser. Report anything it says is missing and how to install it.

## Brand facts to never get wrong
Colours: Washed Black #282829, Cream #f7eddb, Lumber #e0b977 (primary); Olive #978745, Brick #c35d27, Slate #8ba5bc (secondary). No white backgrounds, no gradients, no mixing secondaries. Display type is Apparat (stand-in Familjen Grotesk); body is Inter. Logo is never altered.
