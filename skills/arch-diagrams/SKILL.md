---
name: arch-diagrams
description: Draw editable architecture diagrams into PowerPoint (native zones, icon cards, arrows, step badges) with general, AWS, Azure or GCP icons, using the archdiagrams tool. Use for solution or reference architectures, flows and design comparisons in decks and documents.
---

# Architecture diagrams with archdiagrams

`archdiagrams` turns a short YAML spec into a diagram made of native PowerPoint shapes: dashed zones, white
cards with an icon, a bold name and a small coloured detail line, coloured arrows and numbered badges. The result is
fully editable in PowerPoint and scales into any slide. Use it whenever a slide, document or presentation needs an
architecture picture, a data or event flow, or two designs side by side. Prefer it to an image or a bullet list.

## 1. Make sure the tool is available

```bash
archdiagrams --version || pip install "git+https://github.com/prashant-baj/pptx-archdiagrams"
```

If that install fails (repo not reachable or private), install from a local clone: `pip install -e <path to clone>`.
With pip as root on Linux add `--break-system-packages`. All icons are included: about 150 general ones plus AWS,
Azure, GCP, Kubernetes, on-prem, SaaS and other provider packs. PNG previews need LibreOffice; if the session has its own LibreOffice wrapper (for example the
pptx skill's `scripts/office/soffice.py`), use that to convert the deck to PDF and `pdftoppm` for PNGs.

## 2. Plan before writing YAML

Decide, in this order:
1. **The one message** of the diagram. It becomes the caption.
2. **Tiers, left to right**: who uses it → edge / access → application → data → external systems. Embedded or IoT
   work: device → edge → cloud → people and systems. One zone per tier; 2–5 cards per zone; 12–20 cards in total.
3. **Arrows that matter**: main path solid; control, security or async dashed; two-way only when it really is.
4. **A sequence?** Add badges 1..n on the cards in order, and repeat the sequence in the caption.
5. **Icons**: search with `archdiagrams icons <word>` (`archdiagrams icons general/` lists the general pack,
   `archdiagrams icons aws/database/` one AWS category). Use provider icons when the audience thinks in that cloud,
   general icons otherwise. Do not mix AWS and Azure icons in one diagram unless the design is multi-cloud.

## 3. Write the spec

Start from a starter: `archdiagrams new spec.yaml --kind tiers|flow|compare`. Coordinates are inches inside `size`
(14 x 5.6 suits a 16:9 slide under a title). Key patterns:

```yaml
theme: default                       # default | navy | aws | azure | mono | theme file
size: [14, 5.6]
slide: {title: ..., eyebrow: ..., caption: ..., notes: ...}
elements:
  - zone: {id: site, at: [0, 0, 3, 5.6], label: Site · devices, color: navy, align: spread,
           cards: [{id: ups, icon: ups, title: UPS, sub: SNMP}, {id: mtr, icon: meter, title: Meters, sub: Modbus}]}
  - card: {id: gw, at: [3.4, 0.5, 2.5, 1.0], icon: gateway, title: Site gateway, sub: store and forward, color: teal}
  - bus: {from: [ups, mtr], to: gw, color: navy}            # many-to-one
  - arrow: {from: gw.r, to: db.l, color: teal, label: events}  # anchors: id.l/r/t/b[@0.25]
  - arrow: {from: a.b, to: b.t, route: straight, dashed: true, both: true}
  - badges: [[ups, 1], [gw, 2], [db, 3]]
  - bar: {at: [0, 5.0, 14, 0.6], text: One data model on every link}
  - note: {at: [0, 5.0, 6], text: "Think about: what if the link is down?", color: orange, bold: true, italic: false}
```

Rules that avoid rework:
* Quote any text containing `?`, `:`, `#` or `,` inside `{...}`.
* Card titles 1–3 words; details go in `sub`. Cards about 1.0 in tall and at least 2.0 in wide with an icon.
* Leave at least 0.3 in between cards and 0.8 in between zones where an arrow carries a label; otherwise drop the
  label or set `label_at: [x, y]`.
* Colour by role: one colour per zone; arrows take the colour of their destination tier.

## 4. Render

* New deck: `archdiagrams render spec.yaml -o out.pptx` (several specs → one slide each; `--template company.pptx
  --layout "Title Only"` to use a company template).
* Existing deck, new slide: `archdiagrams add spec.yaml --deck deck.pptx --layout "Title Only" -o new.pptx`
* Existing slide: `archdiagrams add spec.yaml --deck deck.pptx --slide 7 --box 0.9,2.3,14.2,5.6 -o new.pptx`
  (box = the free area of that slide, in inches). Never overwrite the user's original deck; write a new file.
* From Python (inside a larger python-pptx build): `Diagram(slide, size=(14, 5.6), box=(x, y, w, h), theme=...)`,
  then `.zone() .card() .arrow() .bus() .badge() .bar() .note()`.

## 5. Check before you hand it over

1. Read every warning the tool prints (overlaps, text that may not fit, labels with no room) and fix the spec.
2. Render to PNG and look at it: labels on lines, arrows crossing cards, cramped or empty areas, unequal gaps.
3. Fix, re-render, look again. Usually one or two rounds.
4. Tell the user the diagram is editable in PowerPoint, and keep the YAML next to the deck so it can be rebuilt.
