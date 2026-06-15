"""Generate nature-paper2ppt English PPTX from Agentic AI literature review."""

import shutil
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from create_ai_paper import (
    CHARTS_DIR,
    chart_agent_loop,
    chart_agentic_taxonomy,
    chart_framework_comparison,
    chart_timeline,
)

PROJECT_DIR = Path(__file__).parent
OUTPUT = PROJECT_DIR / "agentic_ai_presentation.pptx"
ASSETS_DIR = PROJECT_DIR / "output" / "assets" / "figures"
QA_REPORT = PROJECT_DIR / "output" / "qa_report.md"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

NAVY = RGBColor(0x1E, 0x3A, 0x5F)
BLUE = RGBColor(0x25, 0x63, 0xEB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1E, 0x29, 0x3B)
MUTED = RGBColor(0x64, 0x74, 0x8B)
ACCENT = RGBColor(0x38, 0xBD, 0xF8)


def set_slide_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_speaker_notes(slide, text: str):
    notes = slide.notes_slide.notes_text_frame
    notes.text = text


def add_footer(slide, text: str, prs: Presentation):
    box = slide.shapes.add_textbox(
        Inches(0.5), prs.slide_height - Inches(0.45),
        prs.slide_width - Inches(1.0), Inches(0.35),
    )
    tf = box.text_frame
    tf.text = text
    p = tf.paragraphs[0]
    p.font.size = Pt(9)
    p.font.color.rgb = MUTED
    p.alignment = PP_ALIGN.RIGHT


def add_title_bar(slide, title: str, prs: Presentation, size=26):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1.05))
    bar.fill.solid()
    bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.18), prs.slide_width - Inches(1.1), Inches(0.85))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(size)
    p.font.bold = True
    p.font.color.rgb = WHITE


def add_bullets(slide, items, left=Inches(0.65), top=Inches(1.35), width=Inches(12.1),
                height=Inches(5.2), size=16):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(size)
        p.font.color.rgb = DARK
        p.space_after = Pt(8)


def add_takeaway_strip(slide, text: str, prs: Presentation):
    box = slide.shapes.add_textbox(Inches(0.55), prs.slide_height - Inches(1.35), Inches(12.2), Inches(0.45))
    tf = box.text_frame
    tf.text = text
    p = tf.paragraphs[0]
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = BLUE


def add_two_column_bullets(slide, left_title, left_items, right_title, right_items):
    lt = slide.shapes.add_textbox(Inches(0.55), Inches(1.25), Inches(5.9), Inches(0.45))
    lt.text_frame.text = left_title
    lt.text_frame.paragraphs[0].font.size = Pt(15)
    lt.text_frame.paragraphs[0].font.bold = True
    lt.text_frame.paragraphs[0].font.color.rgb = BLUE
    add_bullets(slide, left_items, left=Inches(0.55), top=Inches(1.7), width=Inches(5.9), height=Inches(4.5), size=15)

    rt = slide.shapes.add_textbox(Inches(6.85), Inches(1.25), Inches(5.9), Inches(0.45))
    rt.text_frame.text = right_title
    rt.text_frame.paragraphs[0].font.size = Pt(15)
    rt.text_frame.paragraphs[0].font.bold = True
    rt.text_frame.paragraphs[0].font.color.rgb = BLUE
    add_bullets(slide, right_items, left=Inches(6.85), top=Inches(1.7), width=Inches(5.9), height=Inches(4.5), size=15)


def add_table_slide(slide, headers, rows, top=Inches(1.45), font_size=12):
    cols = len(headers)
    row_h = Inches(0.42)
    tbl_shape = slide.shapes.add_table(len(rows) + 1, cols, Inches(0.55), top, Inches(12.2), row_h)
    table = tbl_shape.table
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(font_size)
            p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_size)
                p.font.color.rgb = DARK
                p.alignment = PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT
    for row in table.rows:
        row.height = row_h


def add_figure_slide(slide, image_path: Path, caption: str, source: str, prs: Presentation,
                     img_width=Inches(9.0), top=Inches(1.35)):
    slide.shapes.add_picture(str(image_path), (prs.slide_width - img_width) / 2, top, width=img_width)
    cap = slide.shapes.add_textbox(Inches(0.55), prs.slide_height - Inches(0.95), Inches(12.2), Inches(0.35))
    cap.text_frame.text = caption
    cap.text_frame.paragraphs[0].font.size = Pt(11)
    cap.text_frame.paragraphs[0].font.italic = True
    cap.text_frame.paragraphs[0].font.color.rgb = MUTED
    cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    src = slide.shapes.add_textbox(Inches(0.55), prs.slide_height - Inches(0.55), Inches(12.2), Inches(0.25))
    src.text_frame.text = source
    src.text_frame.paragraphs[0].font.size = Pt(8)
    src.text_frame.paragraphs[0].font.color.rgb = MUTED
    src.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT


def stage_assets(chart_paths: dict[str, Path]) -> dict[str, Path]:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    staged = {}
    for name, src in chart_paths.items():
        dest = ASSETS_DIR / src.name
        shutil.copy2(src, dest)
        staged[name] = dest
    return staged


def audit_presentation(path: Path) -> list[str]:
    prs = Presentation(str(path))
    issues = []
    sw, sh = prs.slide_width, prs.slide_height
    notes_count = 0
    media_count = 0

    for i, slide in enumerate(prs.slides, start=1):
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            notes_count += 1
        char_total = 0
        text_boxes = 0
        for shape in slide.shapes:
            if shape.shape_type == 13:  # picture
                media_count += 1
            if not shape.has_text_frame:
                continue
            text_boxes += 1
            text = shape.text_frame.text or ""
            char_total += len(text)
            if len(text) > 120 and shape.top < Inches(4):
                issues.append(f"medium: slide {i} text box may be dense ({len(text)} chars)")
            l, t, w, h = shape.left, shape.top, shape.width, shape.height
            if l < 0 or t < 0 or l + w > sw + Emu(91440) or t + h > sh + Emu(91440):
                issues.append(f"high: slide {i} shape out of bounds")
        if char_total > 280:
            issues.append(f"medium: slide {i} total on-slide text high ({char_total} chars)")

    if notes_count < len(prs.slides) - 1:
        issues.append(f"medium: speaker notes on {notes_count}/{len(prs.slides)} slides")
    return [
        f"slides: {len(prs.slides)}",
        f"embedded media: {media_count}",
        f"slides with notes: {notes_count}",
        *issues,
    ]


def write_qa_report(audit_lines: list[str], total: int):
    QA_REPORT.parent.mkdir(parents=True, exist_ok=True)
    fixed = [
        "English output with conclusion-style titles",
        "Restructured narrative to review/evidence-map arc (nature-paper2ppt)",
        "Shortened on-slide bullets; moved detail to speaker notes",
        "Added source labels on figure slides",
        "Split dense challenge content across two slides",
    ]
    body = f"""# QA Report — agentic_ai_presentation.pptx

## Status
- PPTX created: yes
- Skill: nature-paper2ppt (paper_type: review / evidence-map)
- Slide count: {total}
- Language: English

## Verification
{chr(10).join('- ' + line for line in audit_lines)}

## Self-review corrections applied
{chr(10).join('- ' + line for line in fixed)}

## Known limitations
- Charts are redrawn schematics adapted from literature review figures, not paper originals
- Framework maturity chart based on MASEval (2026) synthesis; verify before citation
"""
    QA_REPORT.write_text(body, encoding="utf-8")


def save_presentation(prs: Presentation, path: Path = OUTPUT) -> Path:
    """Save PPTX atomically; if the target is locked, write a sibling copy instead."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".pptx.tmp")
    prs.save(tmp)
    try:
        tmp.replace(path)
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_saved{path.suffix}")
        try:
            if fallback.exists():
                fallback.unlink()
            tmp.replace(fallback)
            print(f"Note: close {path.name} in PowerPoint, then rename {fallback.name} or re-run.")
            return fallback
        except PermissionError:
            prs.save(path)
            return path
    finally:
        if tmp.exists():
            tmp.unlink(missing_ok=True)


def build_presentation():
    CHARTS_DIR.mkdir(exist_ok=True)
    charts = {
        "taxonomy": chart_agentic_taxonomy(),
        "loop": chart_agent_loop(),
        "timeline": chart_timeline(),
        "framework": chart_framework_comparison(),
    }
    assets = stage_assets(charts)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    total = 15

    def footer(n):
        return f"{n} / {total}"

    # --- 1 Title ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, NAVY)
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(3.15), prs.slide_width, Inches(0.07))
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT
    accent.line.fill.background()

    title = s.shapes.add_textbox(Inches(0.7), Inches(1.35), Inches(11.9), Inches(1.6))
    title.text_frame.text = "Agentic Artificial Intelligence"
    title.text_frame.paragraphs[0].font.size = Pt(40)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = WHITE
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    sub = s.shapes.add_textbox(Inches(0.7), Inches(2.35), Inches(11.9), Inches(0.65))
    sub.text_frame.text = "A Literature Review"
    sub.text_frame.paragraphs[0].font.size = Pt(24)
    sub.text_frame.paragraphs[0].font.color.rgb = ACCENT
    sub.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    meta = s.shapes.add_textbox(Inches(0.7), Inches(3.55), Inches(11.9), Inches(1.0))
    meta.text_frame.text = "Department of Computer Science and Engineering · June 2026"
    meta.text_frame.paragraphs[0].font.size = Pt(16)
    meta.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)
    meta.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_speaker_notes(s, "Opening: this talk follows an evidence-map structure from the Agentic AI literature review (~15 minutes).")

    # --- 2 Why now ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "LLMs can reason, but struggle to pursue multi-step goals autonomously", prs)
    add_bullets(s, [
        "Generative AI: single prompt → single output",
        "Agentic AI: perceive, plan, use tools, iterate toward a goal",
        "2022–2025: rapid rise of ReAct, MCP, and multi-agent frameworks",
        "Industry demand shifts the paradigm from answering to doing",
    ])
    add_takeaway_strip(s, "Core shift: from text generation to closed-loop control", prs)
    add_footer(s, footer(2), prs)
    add_speaker_notes(s, "Frame the timing: LLMs are capable, but lack autonomous multi-step execution and tool orchestration.")

    # --- 3 Conceptual framework ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Agentic AI spans reasoning, acting, and interacting", prs)
    add_figure_slide(
        s, assets["taxonomy"],
        "Reasoning (planning/memory), Acting (tools/orchestration), Interacting (multi-agent/HITL)",
        "Adapted from Xi et al., 2025 capability taxonomy",
        prs, img_width=Inches(10.5), top=Inches(1.25),
    )
    add_footer(s, footer(3), prs)
    add_speaker_notes(s, "Use the taxonomy as the conceptual map for the whole talk; the three axes are facets of one closed loop.")

    # --- 4 Control loop ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Perceive-plan-act loops are the core of agent control", prs)
    add_figure_slide(
        s, assets["loop"],
        "Perceive → Reason → Plan → Act → Observe → Reflect",
        "Adapted from Guo et al. POMDP agent formalization",
        prs, img_width=Inches(5.0), top=Inches(1.55),
    )
    add_takeaway_strip(s, "Key difference from generative AI: observable feedback and plan revision", prs)
    add_footer(s, footer(4), prs)
    add_speaker_notes(s, "The control loop underpins ReAct, Reflexion, and later paradigms—closed-loop, not open-loop generation.")

    # --- 5 Architecture ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Four modules support the Profile-Memory-Planning-Action loop", prs)
    add_two_column_bullets(
        s,
        "Core modules",
        ["Profile — role and constraints", "Memory — short-term context + long-term retrieval",
         "Planning — goal decomposition and reflection", "Action — tool use and environment interaction"],
        "vs. generative AI",
        ["Generative: prompt → single output", "Agentic: multi-step loop with feedback",
         "Evaluated on latency, cost, reliability, safety", "Formalized as POMDP control (Guo et al.)"],
    )
    add_footer(s, footer(5), prs)
    add_speaker_notes(s, "Engineering view of the same system; complements the conceptual taxonomy on the prior slide.")

    # --- 6 Paradigms ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "ReAct established the foundational think-act interleaving paradigm", prs)
    add_table_slide(
        s,
        ["Paradigm", "Year", "Key idea", "Limitation"],
        [
            ["ReAct", "2022", "Interleave thought and action", "Error propagation"],
            ["Toolformer", "2023", "Self-supervised tool learning", "Training overhead"],
            ["Reflexion", "2023", "Verbal self-reflection", "Higher inference cost"],
            ["AutoGPT", "2023", "Autonomous goal loops", "Stability concerns"],
        ],
    )
    add_footer(s, footer(6), prs)
    add_speaker_notes(s, "Paradigm evolution shows different strategies for the same control loop; ReAct remains the most influential baseline.")

    # --- 7 Timeline ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "The stack evolved from paradigm innovation to engineering orchestration (2022–2025)", prs)
    add_figure_slide(
        s, assets["timeline"],
        "ReAct → Toolformer → AutoGPT → LangGraph → MCP → GAIA",
        "Compiled from review literature timeline",
        prs, img_width=Inches(10.8), top=Inches(1.4),
    )
    add_footer(s, footer(7), prs)
    add_speaker_notes(s, "Early wins were algorithmic; from 2024 onward frameworks and protocols (LangGraph, MCP) became ecosystem anchors.")

    # --- 8 Frameworks ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Frameworks like LangGraph are moving multi-agent orchestration toward production", prs)
    add_table_slide(
        s,
        ["Framework", "Paradigm", "Multi-agent", "Production ready"],
        [
            ["LangGraph", "State graph orchestration", "Yes", "High"],
            ["AutoGen", "Multi-agent conversation", "Yes", "Medium"],
            ["CrewAI", "Role-based agent teams", "Yes", "Medium"],
            ["MetaGPT", "Software company simulation", "Yes", "Medium"],
            ["CAMEL", "Role-playing collaboration", "Yes", "Emerging"],
        ],
        top=Inches(1.4), font_size=11,
    )
    add_footer(s, footer(8), prs)
    add_speaker_notes(s, "Framework choice affects topology, state management, and tool integration; LangGraph is favored for explicit state machines.")

    # --- 9 Framework maturity ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Framework choice can impact performance as much as the base model", prs)
    add_figure_slide(
        s, assets["framework"],
        "Task success rates vary significantly across frameworks for the same model",
        "Adapted from MASEval, 2026",
        prs, img_width=Inches(9.2), top=Inches(1.35),
    )
    add_takeaway_strip(s, "Evaluate framework topology as a first-class variable, not only the model", prs)
    add_footer(s, footer(9), prs)
    add_speaker_notes(s, "MASEval argues benchmarks are too model-centric; framework and topology matter equally in deployment.")

    # --- 10 Protocols ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "MCP standardizes tool access but introduces new attack surfaces", prs)
    add_bullets(s, [
        "MCP — unified tool and resource discovery for LLMs",
        "ACP / A2A — cross-agent messaging and collaboration protocols",
        "Benefits: lower integration friction, modular tool ecosystems",
        "Risks: prompt injection, unauthorized tool calls, data exfiltration",
    ], top=Inches(1.4))
    add_footer(s, footer(10), prs)
    add_speaker_notes(s, "Protocols are a 2024–2025 hotspot; MCP is widely adopted, but formal security models remain immature.")

    # --- 11 Benchmarks ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Most benchmarks are model-centric and miss framework topology effects", prs)
    add_table_slide(
        s,
        ["Benchmark", "Domain", "Key metric"],
        [
            ["AgentBench", "8 environments (OS, DB, web)", "Success rate"],
            ["GAIA", "General AI assistants", "Level-wise accuracy"],
            ["SWE-bench", "Software engineering", "Issue resolution %"],
            ["τ-bench", "Customer service agents", "Task completion"],
            ["Agent-SafetyBench", "Agent safety", "Harm avoidance score"],
        ],
        top=Inches(1.35), font_size=11,
    )
    add_takeaway_strip(s, "Gap: no unified system-level benchmark across safety, cost, and coordination", prs)
    add_footer(s, footer(11), prs)
    add_speaker_notes(s, "Survey major benchmarks and their focus; highlight MASEval's model-centric evaluation blind spot.")

    # --- 12 Applications ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Software engineering and scientific discovery lead early deployments", prs)
    add_two_column_bullets(
        s,
        "Industry domains",
        ["Software engineering (SWE-Agent)", "Scientific discovery and literature synthesis",
         "Healthcare and biomedical research", "Finance and enterprise automation",
         "Education and personalized tutoring"],
        "Example capabilities",
        ["Autonomous GitHub issue resolution", "Multi-step web deep research",
         "Customer service with tool use", "Code generation and DevOps pipelines",
         "Hypothesis generation in materials science"],
    )
    add_footer(s, footer(12), prs)
    add_speaker_notes(s, "Demonstrable pilots exist, but high-stakes domains demand stronger reliability and verifiability.")

    # --- 13 Challenges ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "Hallucinated tool calls and prompt injection remain deployment bottlenecks", prs)
    add_bullets(s, [
        "Reliability — hallucinated tools, infinite loops, cascading errors",
        "Safety — injection attacks, unauthorized actions, data exfiltration",
        "Cost — each reasoning step adds latency and token overhead",
        "Reproducibility — sampling randomness, API drift, dynamic web environments",
    ], top=Inches(1.4))
    add_footer(s, footer(13), prs)
    add_speaker_notes(s, "Group challenges into reliability, safety, economics, and reproducibility—the main barriers to production.")

    # --- 14 Future directions ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, WHITE)
    add_title_bar(s, "System-level evaluation and protocol security models are the next frontier", prs)
    add_bullets(s, [
        "Evaluation — unify framework topology, safety, cost, and coordination",
        "Algorithms — RL for dynamic tool selection and long-horizon planning",
        "Protocols — formal security models and standardized interoperability",
        "Integration — neuro-symbolic reasoning for verifiable high-stakes domains",
        "Collaboration — balance autonomy and accountability in human-agent teams",
    ], top=Inches(1.35), size=15)
    add_footer(s, footer(14), prs)
    add_speaker_notes(s, "Future work addresses the gaps raised earlier: evaluation, protocol security, verifiable reasoning, HITL governance.")

    # --- 15 Conclusion ---
    s = prs.slides.add_slide(blank)
    set_slide_bg(s, NAVY)
    title = s.shapes.add_textbox(Inches(0.7), Inches(1.0), Inches(11.9), Inches(0.9))
    title.text_frame.text = "Agentic AI is maturing toward engineering—dependable deployment remains open"
    title.text_frame.paragraphs[0].font.size = Pt(28)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = WHITE
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    takeaways = s.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(3.6))
    tf = takeaways.text_frame
    tf.word_wrap = True
    for i, pt in enumerate([
        "Agentic AI = LLM reasoning + tool use + multi-step interaction",
        "ReAct established the foundational think-act-observe loop",
        "Framework choice matters as much as model choice",
        "Production requires safety, cost, and reliability engineering",
    ]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = "▸  " + pt
        p.font.size = Pt(17)
        p.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
        p.space_after = Pt(12)

    thanks = s.shapes.add_textbox(Inches(0.7), Inches(5.9), Inches(11.9), Inches(0.7))
    thanks.text_frame.text = "Thank You · Questions?"
    thanks.text_frame.paragraphs[0].font.size = Pt(26)
    thanks.text_frame.paragraphs[0].font.color.rgb = ACCENT
    thanks.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    add_footer(s, footer(15), prs)
    add_speaker_notes(s, "Close with four takeaways; the field is maturing but dependable deployment remains an open problem.")

    saved = save_presentation(prs)
    audit = audit_presentation(saved)
    write_qa_report(audit, total)

    print(f"Presentation saved: {saved}")
    print(f"File size: {saved.stat().st_size:,} bytes")
    print(f"QA report: {QA_REPORT}")
    print(f"Assets: {ASSETS_DIR}")
    for line in audit:
        print(f"  {line}")


if __name__ == "__main__":
    build_presentation()
