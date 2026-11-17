---
name: animated-pptx-deck
description: Build a "wow" school/student presentation as a real .pptx with Morph transitions, auto-playing entrance animations, chapter cards and a zoom-out finale, in the editorial light style the user liked (paper canvas, serif headlines). Use when the user asks for a beautiful/animated/TikTok-style PowerPoint presentation on any subject, or says "like the matrices / photo presentation".
---

# Animated PPTX deck (the "Every photo is a matrix" style)

Two finished decks built this way, approved by the user:
- `~/Documents/Study/matlab/photo_matrix_presentation/` (light, editorial — **the preferred style**)
- `~/Documents/Study/matlab/matrices_presentation/` (older dark version)

Working copies of the engine are in `reference/` next to this file:
`anim.py` (animation engine), `build_light_photo.py` (full light deck to adapt),
`build_dark_matrices.py`, `speaker_notes_example.md`.

## Workflow

1. **Agree on the plan before building.** Offer 4-6 topic ideas, then a slide-by-slide plot
   (~15 slides, 3-4 min), and wait for "go". The user dictates by voice: read the whole message.
2. **Palette:** the user sends romanuke.com palette links. Read the hex codes from the page
   (`document.body.innerText.match(/#[0-9A-Fa-f]{6}/g)`). Pick ONE palette, do not mix.
   Light theme: canvas = lightest tint (e.g. `FAF6EE`), ink = darkest colour, 1 accent colour
   for key words, 1 soft colour for highlights, a darker grey for small text.
3. **Write `speaker_notes.md` first** (`### Slide N · Title`, then `> English line`, then
   `*Russian meaning*`). Simple English (A2-B1), short sentences, present tense. `build.py`
   parses it and puts the `>` lines into each slide's notes. Mark auto slides "(auto) No words".
4. **Copy `reference/anim.py` + `reference/build_light_photo.py`** into a new project folder
   and rewrite the `build()` slide section for the new topic. Keep the helpers.
5. **Build:** `python3 build.py` (pass 1 renders slide snapshots through PowerPoint PDF export for
   the zoom-out finale; `--skip-snaps` reuses them).
6. **QA** (see below), send the contact sheet with SendUserFile, open the deck, and ask the
   user to watch the slideshow — motion cannot be verified here.
7. When the user asks to "write the notes", reply with plain text "Слайд 1 — …" (not the md file).

## Style rules (what the user liked)

- Light editorial look: paper canvas + faint grain image, lots of air, magazine layout.
- Fonts (free stand-ins for Anthropic's proprietary type, installed in ~/Library/Fonts):
  **Source Serif 4 Display Light** headlines (key word in *italic* + accent colour, letter
  spacing -60), **Inter** body, **JetBrains Mono** numbers/labels (small caps-style labels
  `01 — PART`, spacing 300). Check glyphs exist (Inter has no `→ ● ⇄`; JetBrains Mono has `→ ↔ ●`).
- One visual motif carried through the whole deck (photo prints with white border + soft
  shadow, viewfinder corners, contact sheet for the photo topic). Pick a motif that fits the
  new subject.
- All graphics drawn by code (PIL/numpy or shapes) — never stock icons/emoji, never anything
  that looks like a template or AI generator. No accent bars under titles.
- No "Matrix movie" clichés for maths topics. Max ~6 words of on-slide text; the talk is in notes.
- Ending: no "Thank you"/"Questions?" slide and no audience quiz. Use the **zoom-out**: the last
  scene shrinks into one cell of a 3×3 sheet of earlier slide snapshots, then the title
  headline returns over the faded sheet (closes the loop with slide 1).

## Animation system

- **Morph** (`morph_xml(dur)`) only moves objects that exist on both slides with the same
  `!!name`. Use it for real "wow" moves: numbers flying, zooms into a crop, flips, the zoom-out.
- Nothing appears from nowhere: every new object gets an entrance in `Timeline` that starts
  automatically after the transition (one click per slide). Kinds: `mask` (text rises from
  behind an invisible edge — default for titles), `rise`, `fly` (+`bounce`), `drop`, `pop`,
  `spin`, `grow`, `stretchx/y`, `wipe` (`dir="left"` = reveals left→right!), `wheel`, `slide`
  (linear move, for sliders/scan lines), `pulse`, `out`, `fallout`.
- Old objects leave: previous title parked above the slide (`gone_up`) so Morph flies it out;
  other leftovers go off-screen (x + 7.5 or y + 8).
- Consistent rhythm: stagger 0.1-0.15 s, one direction (new from right/bottom, old to left/top).
- Chapter cards: auto slide (`morph_xml(650, adv=1700-1900)`) — iris circle or slanted stripes
  cover the screen with "PART 02 / Title", then open again.
- Mid-move "card flip": an auto slide with `adv=0` where the picture is squeezed to a thin
  edge, then the next slide expands it mirrored.
- Effects that worked: stack of prints dropping, print "developing" (frame first, picture fades
  in), count-up number, scan line revealing rows, before/after slider riding a wipe, pixels
  falling off leaving numbers, circled favourite frame on the contact sheet.

## Hard-won gotchas

- **Never** export to video (`save as movie`) — it crashed the Mac twice.
- QA = PDF export via AppleScript on a **copy** with a unique name, addressed by
  `presentation "qa_copy.pptx"` (never `active presentation` — that closes the user's open
  deck without saving). Wrap osascript in `perl -e 'alarm 150; exec @ARGV'`. Then `pdftoppm`
  + `ffmpeg tile` contact sheet; `pdffonts` to confirm all fonts are real (no Times fallback).
- Each extra effect on the same shape needs its own `grpId` + `bldP` (anim.py does this) or
  PowerPoint silently drops it (a white "flash" once stuck on screen because of this).
- Never leave a full-screen shape visible in the static slide (e.g. a flash) — in edit mode
  the slide looks broken and it may need two clicks in the show.
- PowerPoint's `wipe(right)` reveals from the right; for left→right use `wipe(left)`.
- Validate: `python3 <pptx skill>/scripts/office/validate.py deck.pptx`.
- Tell the user: enable "Embed fonts" (Preferences → Save) before showing on another computer;
  if the deck is open in PowerPoint, close it **without saving** before reopening a new build.
