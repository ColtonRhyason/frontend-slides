# Shift Projects Brand Preset

Brand guide for generating on-brand Shift Projects presentations. **Read this file in full whenever Shift Projects context is detected.** Everything here comes from the Shift Projects Brand Guidelines (Public Eye Studio, March 2025). Do not infer or guess brand rules.

---

## When to Use This Preset

Default to this preset when any of these are true:

- User asks for a "Shift Projects" or "Shift Accessibility" deck or presentation
- User mentions residential construction, custom home builds, home renovations or accessibility renovations in a Shift context
- User uploads Shift-branded source material (existing decks, logos, photos)
- User identifies themselves as working at or consulting for Shift Projects
- Conversation memory indicates Shift Projects work

When the preset applies, skip Phase 2 style discovery. The brand is decided. Go straight to Phase 3 using `shift-projects-slide-library.html` as the starting template.

---

## Source of Truth

| File | Purpose |
|------|---------|
| `shift-projects-brand.md` (this file) | Rules, voice, layout catalog |
| `shift-projects-slide-library.html` | Pre-built template with all layouts, inline editing, logos embedded |
| `assets/shift-projects/*.svg` | Vector brand marks extracted from the guidelines (`currentColor` fill) |

**Workflow:** copy `shift-projects-slide-library.html`, keep the layouts that fit the content, delete the rest, swap placeholder copy for real content following the voice rules below. Never regenerate the CSS or layouts from scratch.

---

## Brand Platform

- **Purpose:** Built for today, ready for tomorrow.
- **Promise:** We promise to deliver an unexpectedly inspiring construction experience.
- **Mission:** To elevate homes, and the way people live, by prioritizing function, design, and process.
- **Vision:** To be the most trusted and sought-after residential contractor, known for genuinely caring about clients and ensuring they get the greatest value from their investment, including long-lasting results.
- **Tagline:** Homes built for today, ready for tomorrow.
- **Values:** Ownership, Respect, Mutual Trust.
- **Personality:** Dependable, Sincere, Straightforward, Polished, Collaborative. Never Clinical, Luxury, Opulent, Aloof or Rustic.
- **Pillars:** Focus is residential projects. Specialties are New Builds, Home Renovations, Accessibility Renovations.

---

## Colours

| Token | Name | Hex | Role |
|-------|------|-----|------|
| `--sp-black` | Washed Black | `#282829` | Primary. Text, reversed backgrounds, graphics |
| `--sp-cream` | Cream | `#f7eddb` | Primary. Default light background |
| `--sp-lumber` | Lumber | `#e0b977` | Primary. Section dividers, covers, accents |
| `--sp-olive` | Olive | `#978745` | Secondary |
| `--sp-brick` | Brick | `#c35d27` | Secondary |
| `--sp-slate` | Slate | `#8ba5bc` | Secondary |

### Colour Rules

- Primary palette leads every layout. Secondaries add contrast and must never overshadow the primaries.
- **Approved:** black on cream; cream on black (reversed); black on lumber, olive, brick or slate; a single secondary colour used as an overlay band.
- **Never:**
  - White as a background. Use cream instead (white only for practical print reasons).
  - Gradients.
  - Mixing secondary colours with each other on one element.
  - Secondary colours over imagery.
  - Off-palette text colours (no blue, no grey type except secondary captions at reduced opacity).
- Text on any brand colour is Washed Black, except on Washed Black where it is Cream.

---

## Typography

### Primary: Apparat (Book)

Headlines and display type. Large, light, tightly tracked, sentence case. Apparat should always be the first thing a viewer sees. Not on Google Fonts; it is a licensed typeface.

**Fallback for HTML decks:** `Familjen Grotesk` from Google Fonts, falling back to `Inter Tight`. If the user supplies Apparat font files, add `@font-face` rules and put `Apparat` first in `--font-display`.

### Secondary: Inter

- **Inter Bold**, uppercase: minor subheads, eyebrow labels, footer and functional info. Add letter-spacing (`0.12em` for small labels).
- **Inter Regular**: body copy and long paragraphs.

### Hierarchy

| Level | Font | Weight | Case | Notes |
|-------|------|--------|------|-------|
| Display / H1 | Apparat | 400 | Sentence | Line-height ~1.0, tracking -2% |
| Major subhead | Inter | 700 | UPPERCASE | Can sit alongside Apparat as a major header |
| Eyebrow, footer | Inter | 700 | UPPERCASE | ~0.7rem, tracking 0.12em |
| Body | Inter | 400 | Sentence | Line-height 1.35 |

**Email and non-web contexts** use web-safe fonts only: Georgia Bold for headers, Trebuchet MS for subheads and body (Arial and Helvetica also approved).

---

## Logo Usage

Three marks, all in `assets/shift-projects/`:

1. **Wordmark.** Primary and the default. "Shift Projects" in one line, centred or left-aligned with generous clear space. Used for all core communications. The template uses the stacked file for the two-line version; for the single-line primary wordmark use the artwork supplied by Public Eye Studio.
2. **Stacked** (`wordmark-stacked.svg`) and **Icon** (`icon.svg`). Secondary. Icon for small or square formats, stacked for vertical layouts. Use only when the primary wordmark does not fit.
3. **Icon lockup.** Icon plus wordmark side by side. Use sparingly where brand familiarity is limited.

Colour: Washed Black on light backgrounds, Cream on Washed Black. No other colours.

**Never** skew, add effects, outline, remake in another font, change the icon/wordmark relationship, crop into a shape, recolour, change alignment, or change proportions.

Placement in the template: the icon sits top-right on every slide by default (press **L** in the HTML deck to hide it, or add `no-logo` to a slide; in PPTX use the "(no logo)" layouts).

Placement detail: the icon sits top-right on cover and content slides; the stacked wordmark is used on the cover and closing.

---

## Supporting Graphics

- **Linear element** (`linear-large.svg`): stepped horizontal bands that echo the icon. Use sparingly for visual interest on covers, dividers, or as a background band. Washed Black on slate or cream; a single secondary colour may overlay as a flat band.
- **Cropping shape:** images and panels get the stepped corner cut (a rounded notch in the lower edge) as a branded crop. The template implements it as `.sp-crop` with an SVG `clip-path`.
- **Photography:** two styles. *Process* shots are behind-the-scenes, documentary, single-colour (warm monochrome sepia). *Project* shots are finished work in full colour, polished. Keep each style consistent within a slide.

---

## Voice and Tone

Tone dials: leans slightly formal over casual, slightly serious over fun, idealistic over realistic, nontraditional over traditional, and more professional than friendly. Approachable and exclusive sit in the middle.

1. Use clear, direct language.
2. Prioritize understanding and "people-first" framing.
3. Communicate process and quality with confidence.
4. Avoid jargon, aloofness, and overly luxurious descriptors.

Copy examples: "People-first home builders." "Your home. Your perspective." Headlines are short, sentence case, usually ending in a full stop on statement slides.

---

## Layout Catalog

Implemented in `shift-projects-slide-library.html` (19 layouts). Frame on every slide: eyebrow row at top (section label, short rule, subsection), footer row at bottom (SHIFT PROJECTS, deck title, page number, filled automatically).

| # | Class | Background | Use |
|---|-------|-----------|-----|
| 1 | `cover` | Lumber | Title slide with the linear graphic sliding in |
| 2 | `divider` | Lumber | Section opener: huge section name, number top-right |
| 3 | `agenda` | Cream | Contents: lead-in left, numbered section list |
| 4 | `split` | Cream | Headline left, body copy right |
| 5 | `numbered` | Cream | 3-4 numbered columns |
| 6 | `statement` | Slate | One big tagline-style statement plus a small paragraph |
| 7 | `reversed` | Washed Black | Key figure or claim, cream type, brick bands that grow in |
| 8 | `stats` | Cream | Three big numbers with rules, label and one-line context |
| 9 | `feature` | Cream + Olive | Project or case study: text and bullets left, branded-crop panel with photo and icon right |
| 10 | `photos` | Cream | Photo grid: process (sepia) left, project (full colour) right |
| 11 | `quote` | Olive | Testimonial in large display type |
| 12 | `dials` | Cream | Spectrum rows with brand-colour dots that settle on their value |
| 13 | `pillars` | Cream | Focus and specialty hierarchy with rules |
| 14 | `dodont` | Cream | Approved vs unapproved comparison, crossed-out samples |
| 15 | `cards` | Cream | Four colour-block cards (black, cream, lumber, slate) |
| 16 | `steps` | Cream | Four stepped black bars echoing the linear graphic: process or timeline |
| 17 | `poll` | Lumber | Audience question with lettered options |
| 18 | `recap` | Cream | Three-column outline with numbered topics |
| 19 | `closing` | Lumber | Thank you, contact, icon plus stacked wordmark |

Pacing: one cover, dividers between sections, mostly cream content slides, and a coloured slide (statement, quote, reversed, poll) every three or four slides for rhythm.

**Photos:** every `.photo` is a placeholder. Swap in `<div class="photo process"><img src="..."></div>` (auto-converted to single-colour sepia) or `<div class="photo project"><img src="..."></div>` (full colour).

**Motion:** content fades up on arrival; bands and bars grow in; the linear graphic slides in; tone-dial dots travel to their value. All of it is disabled under `prefers-reduced-motion`.
