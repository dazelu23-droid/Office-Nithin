---
name: nature-skills
description: >-
  Nature-family academic writing and research toolkit bundle (Yuan1z0825/nature-skills).
  Routes to specialized sub-skills for figures, polishing, writing, peer review, citations,
  data availability, paper reading, reviewer responses, paper-to-PPT, paper-to-patent, and
  academic search. Use when the user mentions Nature style, academic writing, scientific figures,
  paper translation, journal club PPT, reviewer response, patent drafting, literature search,
  or Chinese triggers like 论文润色、科研绘图、文献翻译、组会PPT、审稿意见回复、论文转专利.
---

# Nature Skills — Router

This project bundles [nature-skills](https://github.com/Yuan1z0825/nature-skills) as project-level Cursor skills under `.cursor/skills/`. Each sub-skill is self-contained; many depend on shared content in `_shared/`.

**Do not substitute generic advice.** When a sub-skill matches, read its `SKILL.md` and follow its routing protocol (load `manifest.yaml`, `static/`, and `references/` from disk as directed).

## Pick a sub-skill

| Sub-skill | Use when | Path |
|-----------|----------|------|
| `nature-figure` | Publication plots, multi-panel figures, matplotlib/ggplot, SVG export | [../nature-figure/SKILL.md](../nature-figure/SKILL.md) |
| `nature-polishing` | Polish/translate prose to Nature-leaning English; LaTeX layout fixes | [../nature-polishing/SKILL.md](../nature-polishing/SKILL.md) |
| `nature-writing` | Draft abstracts, intros, results, discussion, methods, outlines | [../nature-writing/SKILL.md](../nature-writing/SKILL.md) |
| `nature-reviewer` | Pre-submission reviewer simulation (3 reports + synthesis) | [../nature-reviewer/SKILL.md](../nature-reviewer/SKILL.md) |
| `nature-citation` | Nature/CNS-family citation search; ENW/RIS/Zotero RDF export | [../nature-citation/SKILL.md](../nature-citation/SKILL.md) |
| `nature-data` | Data Availability statements, repository plans, FAIR checks | [../nature-data/SKILL.md](../nature-data/SKILL.md) |
| `nature-reader` | Full bilingual Markdown reader with figure/table grounding | [../nature-reader/SKILL.md](../nature-reader/SKILL.md) |
| `nature-response` | Point-by-point reviewer response / rebuttal letters | [../nature-response/SKILL.md](../nature-response/SKILL.md) |
| `nature-paper2ppt` | Chinese PPTX from papers (journal club, lab meeting) | [../nature-paper2ppt/SKILL.md](../nature-paper2ppt/SKILL.md) |
| `nature-paper-to-patent` | Chinese invention patent drafts from research papers | [../nature-paper-to-patent/SKILL.md](../nature-paper-to-patent/SKILL.md) |
| `nature-academic-search` | Multi-source literature search (PubMed, CrossRef, arXiv) | [../nature-academic-search/SKILL.md](../nature-academic-search/SKILL.md) |

## Invocation protocol

1. Match the user request to one row above (or combine skills only when the user asks for multiple deliverables).
2. Read that sub-skill's `SKILL.md` from the path in the table.
3. Preserve sibling layout: sub-skills reference `../_shared/` — do not move or rename `_shared`.
4. For router-style skills, always load `manifest.yaml` and only the `static/` fragments the manifest selects; pull `references/` on demand.
5. Deliver the sub-skill's primary output (prose, `.svg`, `.pptx`, `.docx`, reference file, etc.) — not a planning memo unless the user asks.

## Shared support

Cross-skill files live in [../_shared/](../_shared/) (terminology ledger, ethics, journal formats, paper-type taxonomy). Load specific files only when a sub-skill points to them.

## Updating

To refresh from upstream:

```powershell
# From project root — re-download and recopy skills
$zip = ".cursor/nature-skills.zip"
Invoke-WebRequest -Uri "https://github.com/Yuan1z0825/nature-skills/archive/refs/heads/main.zip" -OutFile $zip
Expand-Archive -Path $zip -DestinationPath ".cursor/nature-skills-src" -Force
$src = ".cursor/nature-skills-src/nature-skills-main/skills"
Copy-Item "$src/_shared" ".cursor/skills/_shared" -Recurse -Force
Get-ChildItem "$src/nature-*" -Directory | Copy-Item -Destination ".cursor/skills" -Recurse -Force
```

License: MIT — see upstream repository.
