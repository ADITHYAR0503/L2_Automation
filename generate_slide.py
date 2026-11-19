import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

def build_executive_slide():
    prs = Presentation()
    # Set Widescreen 16:9 Dimensions (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Color Palette (Obsidian Black, Carbon Surface, Electric Purple, Crisp White)
    BG_OBSIDIAN = RGBColor(9, 9, 11)
    CARD_SURFACE = RGBColor(18, 18, 26)
    ROYAL_PURPLE = RGBColor(124, 58, 237)
    VIOLET_GLOW = RGBColor(168, 85, 247)
    PURE_WHITE = RGBColor(255, 255, 255)
    MUTED_GRAY = RGBColor(161, 161, 170)

    # Set Slide Background to Obsidian Black
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_OBSIDIAN

    # Top Accent Bar
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ROYAL_PURPLE
    top_bar.line.fill.background()

    # Header Title & Subtitle
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12.3), Inches(0.9))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True

    p_title = tf_title.paragraphs[0]
    p_title.text = "L2 AUTOMATION  //  PRESTO OPERATIONAL COMMAND SUITE"
    p_title.font.name = "Calibri"
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = PURE_WHITE

    p_sub = tf_title.add_paragraph()
    p_sub.text = "Centralizing Daily PCM Reporting, Live L2 Alerting & DR2 Operational Resilience   |   Architected by Adithya"
    p_sub.font.name = "Calibri"
    p_sub.font.size = Pt(12)
    p_sub.font.color.rgb = VIOLET_GLOW

    # Helper Function for Left Cards
    def add_info_card(top_in, height_in, heading, bullets):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(top_in), Inches(6.2), Inches(height_in)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_SURFACE
        card.line.color.rgb = ROYAL_PURPLE
        card.line.width = Pt(1.2)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.18)
        tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.10)
        tf.margin_bottom = Inches(0.08)

        p_head = tf.paragraphs[0]
        p_head.text = heading
        p_head.font.name = "Calibri"
        p_head.font.size = Pt(12)
        p_head.font.bold = True
        p_head.font.color.rgb = VIOLET_GLOW

        for b_title, b_text in bullets:
            p = tf.add_paragraph()
            p.space_before = Pt(3)
            run_bold = p.add_run()
            run_bold.text = f"• {b_title}: "
            run_bold.font.name = "Calibri"
            run_bold.font.size = Pt(10.5)
            run_bold.font.bold = True
            run_bold.font.color.rgb = PURE_WHITE

            run_body = p.add_run()
            run_body.text = b_text
            run_body.font.name = "Calibri"
            run_body.font.size = Pt(10.5)
            run_body.font.color.rgb = MUTED_GRAY

    # Card 1: The Idea & Phase 1 Progress
    add_info_card(
        1.25, 1.55,
        "1. THE IDEA & CURRENT PROGRESS (PHASE 1 LIVE)",
        [
            ("Unified Command Hub", "Replaces manual spreadsheets & scattered tabs with one internal web app."),
            ("PCM Reports Engine", "Computes P1–P3 SLAs & MTTR; generates 1-click unsent Outlook drafts (.eml)."),
            ("Alerts & SOTI Health", "Live L2 alert feed + SOTI Deployment Server connection & queue depth monitor.")
        ]
    )

    # Card 2: Operational Impact
    add_info_card(
        2.92, 1.35,
        "2. WHY THIS MATTERS (TEAM IMPACT)",
        [
            ("95% Time Reduction", "Cuts daily report prep from ~45–60 mins to <2 mins with zero formula errors."),
            ("Human-in-the-Loop Safety", "Never auto-sends emails; analysts review every .eml draft before dispatch.")
        ]
    )

    # Card 3: Outlook Outage Resilience & DR2 Roadmap
    add_info_card(
        4.39, 1.55,
        "3. OUTAGE RESILIENCE & DR2 EXPANSION",
        [
            ("Outlook Outage Mitigation", "Decouples reporting & live alert visibility from email availability during outages."),
            ("DR2 Reports Integration", "Expanding suite to automate DR2 failover verification & secondary reporting."),
            ("Team-Wide Rollout", "Modular Python architecture ready for seamless adoption across all L2 shifts.")
        ]
    )

    # Card 4: Next Steps & Support Needed
    add_info_card(
        6.06, 1.05,
        "4. NEXT STEPS & SUPPORT REQUESTED",
        [
            ("Phase 2 Backend", "Wiring live Excel processing engine to Generate Report + aligning on DR2 templates.")
        ]
    )

    # Right Side Visual Showcase Frame
    frame = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.95), Inches(1.25), Inches(5.88), Inches(5.86)
    )
    frame.fill.solid()
    frame.fill.fore_color.rgb = CARD_SURFACE
    frame.line.color.rgb = ROYAL_PURPLE
    frame.line.width = Pt(1.5)

    img_path = "ui_preview.png"
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(7.08), Inches(1.42), width=Inches(5.62))
    else:
        tf_ph = frame.text_frame
        p_ph = tf_ph.paragraphs[0]
        p_ph.alignment = PP_ALIGN.CENTER
        p_ph.text = "\n\n\n[ Place your UI screenshot as 'ui_preview.png' in C:\\L2_Automation and re-run ]"
        p_ph.font.size = Pt(12)
        p_ph.font.color.rgb = MUTED_GRAY

    # Caption at Bottom of Right Frame
    cap_box = slide.shapes.add_textbox(Inches(7.08), Inches(5.35), Inches(5.62), Inches(1.6))
    tf_cap = cap_box.text_frame
    tf_cap.word_wrap = True
    p_cap1 = tf_cap.paragraphs[0]
    p_cap1.text = "LIVE v1.0 INTERFACE PREVIEW"
    p_cap1.font.name = "Calibri"
    p_cap1.font.size = Pt(11)
    p_cap1.font.bold = True
    p_cap1.font.color.rgb = VIOLET_GLOW

    p_cap2 = tf_cap.add_paragraph()
    p_cap2.space_before = Pt(4)
    p_cap2.text = (
        "• Automated 905 & TTC MTD SLA Summary cards & P1–P3 adherence tables\n"
        "• Interactive MTTR bar visualization by device type\n"
        "• 1-Click Outlook Draft (.eml) generator with pre-populated headers\n"
        "• Real-time PRESTO Ops L2 Alerts & SOTI Deployment Server queue health"
    )
    p_cap2.font.name = "Calibri"
    p_cap2.font.size = Pt(10)
    p_cap2.font.color.rgb = PURE_WHITE

    output_file = "L2_Automation_Overview.pptx"
    prs.save(output_file)
    print(f"SUCCESS: Generated {output_file} in {os.getcwd()}")

if __name__ == "__main__":
    build_executive_slide()