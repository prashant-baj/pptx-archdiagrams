# Changelog

## 0.1.0 - 2026-11-26
First release.
- Native, editable PowerPoint diagrams: zones, icon cards, bars, notes, arrows (right-angle routing, labels,
  dashed, two-way), bus connectors, numbered badges.
- YAML / JSON specs; Python API; one diagram canvas scaled into any slide box.
- CLI: `render`, `add` (onto an existing slide or a new slide from a template layout), `icons` (search and
  contact sheets), `new` (starter specs), `themes`; PNG export through LibreOffice.
- 149 tintable general icons (Tabler / Material); about 2,450 provider icons (AWS, Azure, GCP, Kubernetes, ...) bundled in the package.
- Themes: default, navy, aws, azure, mono, or your own file.
- Checks: overlapping cards, text that may not fit, arrow labels with no room.
- Claude skill in `skills/arch-diagrams`.
