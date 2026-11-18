import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_shading(cell, color_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'    <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'    <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'    <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'    <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'    <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'    <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'    <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'    <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'    <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'    <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'    <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'    <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def create_document():
    doc = Document()

    # Section 0: Preliminary Pages (Cover, Certificate, Declaration, Acknowledgement, Synopsis, TOC)
    sec_prelim = doc.sections[0]
    sec_prelim.top_margin = Inches(1.0)
    sec_prelim.bottom_margin = Inches(1.0)
    sec_prelim.left_margin = Inches(1.25)
    sec_prelim.right_margin = Inches(1.0)
    sec_prelim.page_width = Inches(8.5)
    sec_prelim.page_height = Inches(11.0)

    # Base Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(6)

    # Helper function for adding paragraphs
    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, size=12, space_before=0, space_after=6, line_spacing=1.5):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.italic = italic
        return p

    def add_chapter_heading(num_str, title_str):
        doc.add_page_break()
        p1 = add_p(f"CHAPTER {num_str}", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=16, space_after=4)
        p2 = add_p(title_str.upper(), align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=0, space_after=18)
        return p1, p2

    def add_section_heading(num_title_str):
        p = add_p(num_title_str, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, size=13, space_before=14, space_after=6)
        return p

    def add_subsection_heading(num_title_str):
        p = add_p(num_title_str, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, size=13, space_before=10, space_after=4)
        return p

    # =========================================================================
    # 1. COVER PAGE (Verified against Sri Dharani & 23MSS011)
    # =========================================================================
    add_p("RELEASE INCIDENT CORRELATOR", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=16, space_after=6, line_spacing=1.2)
    add_p("MAJOR PROJECT - I", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=0, space_after=18, line_spacing=1.2)

    add_p("Submitted by", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=11, space_before=0, space_after=2)
    add_p("KAVINRAGHUL V", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("23MSS020", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=0, space_after=18)

    add_p("Under the Guidance of", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=11, space_before=0, space_after=2)
    add_p("Dr. G. S. KARTHICK, M.Sc., DST Fellow (IND)., NET., Ph.D.", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("Assistant Professor", align=WD_ALIGN_PARAGRAPH.CENTER, size=11, space_before=0, space_after=2)
    add_p("Department of Software Systems", align=WD_ALIGN_PARAGRAPH.CENTER, size=11, space_before=0, space_after=18)

    add_p("In partial fulfillment of the requirements for the award of the degree of", align=WD_ALIGN_PARAGRAPH.CENTER, size=11, space_before=0, space_after=2)
    add_p("MASTER OF SCIENCE IN SOFTWARE SYSTEMS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("(Five-Years Integrated Course)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=11, space_before=0, space_after=14)

    # College Emblem (Exact 1.50" x 1.67" proportion matching Sri Dharani P13)
    if os.path.exists("images/psg_college_logo.jpeg"):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(12)
        run_img = p_img.add_run()
        run_img.add_picture("images/psg_college_logo.jpeg", width=Inches(1.50))

    add_p("DEPARTMENT OF SOFTWARE SYSTEMS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=0, space_after=2)
    add_p("PSG COLLEGE OF ARTS & SCIENCE", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("An Autonomous College - Affiliated to Bharathiar University", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Accredited with 'A++' Grade by NAAC (4th Cycle)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("College with Potential for Excellence (Status Awarded by the UGC)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Star College Status Awarded by DBT - MST", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("An ISO 9001:2015 Certified Institution", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Coimbatore - 641 014", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=8)
    add_p("October 2026", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=0, space_after=0)

    # =========================================================================
    # 2. CERTIFICATE PAGE (Verified against Sri Dharani & 23MSS011)
    # =========================================================================
    doc.add_page_break()
    add_p("DEPARTMENT OF SOFTWARE SYSTEMS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=12, space_before=8, space_after=2)
    add_p("PSG COLLEGE OF ARTS & SCIENCE", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=2)
    add_p("An Autonomous College - Affiliated to Bharathiar University", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Accredited with 'A++' Grade by NAAC (4th Cycle)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("College with Potential for Excellence (Status Awarded by the UGC)", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Star College Status Awarded by DBT - MST", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("An ISO 9001:2015 Certified Institution", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=1)
    add_p("Coimbatore - 641 014", align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=0, space_after=18)

    add_p("CERTIFICATE", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=8, space_after=18)

    add_p(
        "This is to certify that this project work entitled “RELEASE INCIDENT CORRELATOR” is a bonafide record of work done by "
        "KAVINRAGHUL V (23MSS020) in partial fulfillment of the requirements for the award of the degree of Master of Science in "
        "Software Systems (Five-Years Integrated Course) of Bharathiar University.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=8, space_after=40, line_spacing=1.5
    )

    # Signatures table: Faculty Guide & Head of the Department
    table_sig = doc.add_table(rows=1, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sig.autofit = False
    set_table_no_borders(table_sig)
    table_sig.cell(0, 0).width = Inches(3.1)
    table_sig.cell(0, 1).width = Inches(3.1)

    c0 = table_sig.cell(0, 0).paragraphs[0]
    c0.paragraph_format.line_spacing = 1.2
    c0.add_run("Faculty Guide\n(Dr. G. S. Karthick)\nAssistant Professor").font.name = 'Times New Roman'

    c1 = table_sig.cell(0, 1).paragraphs[0]
    c1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c1.paragraph_format.line_spacing = 1.2
    c1.add_run("Head of the Department\n(Dr. K. V. Rukmani)\nDepartment of Software Systems").font.name = 'Times New Roman'

    add_p("Submitted for Viva-Voce Examination held on ____________________", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=40, space_after=40)

    # Examiners signature table
    table_exam = doc.add_table(rows=1, cols=2)
    table_exam.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_exam.autofit = False
    set_table_no_borders(table_exam)
    table_exam.cell(0, 0).width = Inches(3.1)
    table_exam.cell(0, 1).width = Inches(3.1)

    e0 = table_exam.cell(0, 0).paragraphs[0]
    e0_run = e0.add_run("Internal Examiner")
    e0_run.font.bold = True
    e0_run.font.name = 'Times New Roman'

    e1 = table_exam.cell(0, 1).paragraphs[0]
    e1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    e1_run = e1.add_run("External Examiner")
    e1_run.font.bold = True
    e1_run.font.name = 'Times New Roman'

    # =========================================================================
    # 3. DECLARATION PAGE (Verified against Sri Dharani & 23MSS011)
    # =========================================================================
    doc.add_page_break()
    add_p("DECLARATION", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=24)

    add_p(
        "I, KAVINRAGHUL V (23MSS020), hereby declare that this project work entitled “RELEASE INCIDENT CORRELATOR” is submitted "
        "to PSG College of Arts & Science (Autonomous), Coimbatore in partial fulfillment for the award of Master of Science in "
        "Software Systems, is a record of original work done by me under the supervision and guidance of Dr. G. S. Karthick, M.Sc., "
        "DST Fellow (IND)., NET., Ph.D., Assistant Professor, Department of Software Systems, PSG College of Arts & Science, Coimbatore.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=10, space_after=14, line_spacing=1.5
    )

    add_p(
        "This project work has not been submitted by me for the award of any other Degree / Diploma / Associateship / Fellowship "
        "or any other similar degree to any other university.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=45, line_spacing=1.5
    )

    table_dec = doc.add_table(rows=1, cols=2)
    table_dec.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_dec.autofit = False
    set_table_no_borders(table_dec)
    table_dec.cell(0, 0).width = Inches(3.1)
    table_dec.cell(0, 1).width = Inches(3.1)

    d0 = table_dec.cell(0, 0).paragraphs[0]
    d0.paragraph_format.line_spacing = 1.2
    d0.add_run("Place : Coimbatore\nDate  : ").font.name = 'Times New Roman'

    d1 = table_dec.cell(0, 1).paragraphs[0]
    d1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d1.paragraph_format.line_spacing = 1.2
    run_d1_name = d1.add_run("KAVINRAGHUL V\n")
    run_d1_name.font.bold = True
    run_d1_name.font.name = 'Times New Roman'
    run_d1_roll = d1.add_run("23MSS020")
    run_d1_roll.font.bold = True
    run_d1_roll.font.name = 'Times New Roman'

    # =========================================================================
    # 4. ACKNOWLEDGEMENT PAGE (Verbatim Dignitaries from 23MSS011 & Sri Dharani)
    # =========================================================================
    doc.add_page_break()
    add_p("ACKNOWLEDGEMENT", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=18)

    add_p(
        "My venture would remain incomplete without expressing my heartfelt gratitude to all those who have contributed to the "
        "successful completion of my project work.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "I would like to thank Thiru L. Gopalakrishnan, Managing Trustee, PSG & Sons' Charities, for providing me with a prospect "
        "and surroundings that made the work possible.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "I take this opportunity to express my deep sense of gratitude to Dr. T. Kannaian, Secretary of PSG College of Arts & Science, "
        "Coimbatore for permitting and doing the needful towards the successful completion of this project.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "I thank our Principal, Dr. M. Jayanthi for her support and constant source of inspiration through the course of the project.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "I am very thankful to the Vice Principal, Dr. M. Umarani for her constant support and motivation throughout the course of this project.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "My heartfelt thanks to Dr. K. V. Rukmani, Head of the Department, Department of Software Systems, for her wholehearted support and "
        "valuable suggestions that greatly contributed to the successful completion of my project.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "I convey my heartiest and most sincere gratitude to my project guide Dr. G. S. Karthick, Assistant Professor, Department of "
        "Software Systems, for his timely suggestions and continuous guidance, which enabled me to complete the project successfully.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing=1.5, space_after=8
    )

    add_p(
        "Finally, this note of acknowledgment would be incomplete without expressing my deepest gratitude to my parents, friends, "
        "and well-wishers for their blessings, encouragement, financial support, and patience, without which the successful completion "
        "of this project would not have been possible.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=24, line_spacing=1.5
    )

    p_ack_sig = add_p("KAVINRAGHUL V\n23MSS020", align=WD_ALIGN_PARAGRAPH.RIGHT, bold=True, space_before=6, space_after=0)

    # =========================================================================
    # 5. SYNOPSIS PAGE (Verbatim Approved Abstract from PDF)
    # =========================================================================
    doc.add_page_break()
    add_p("SYNOPSIS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=8)
    add_p("RELEASE INCIDENT CORRELATOR (RIC)", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=16)

    add_p(
        "In modern continuous integration and continuous deployment (CI/CD) environments, software updates are released frequently "
        "across multi-service distributed architectures. While this rapid delivery accelerates feature delivery, it also introduces operational "
        "risks where code changes trigger downstream regressions, latency spikes, or systemic service degradation. Diagnosing these incidents "
        "is often impeded by fragmented telemetry, siloed release logs, and lack of temporal correlation between deployments and system anomalies. "
        "This project presents Release Incident Correlator (RIC), a unified operational analytics and incident-resolution platform designed to "
        "automate root-cause identification and accelerate Mean Time to Resolution (MTTR). RIC ingests structured release metadata, commit diffs, "
        "telemetry metrics, and incident alerts into a centralized event-correlation pipeline. Leveraging semantic vector representations and "
        "temporal proximity modeling, the system dynamically maps incoming production incidents against recent software releases, identifying "
        "candidate culprit deployments with high confidence.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=10, line_spacing=1.5
    )

    add_p(
        "Furthermore, the platform integrates an automated mitigation orchestration engine that generates targeted rollback plans, canary verification "
        "strategies, and incident remediation runbooks. Developed with a high-performance backend, robust relational persistence with vector extensions, "
        "and an interactive operational dashboard, RIC provides engineering teams with real-time incident scoring, commit-level culpability attribution, "
        "and actionable resolution intelligence. By bridging the gap between deployment pipelines and incident response workflows, RIC enhances system "
        "reliability, minimizes downtime, and supports scalable observability in enterprise cloud environments.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=10, line_spacing=1.5
    )

    add_p(
        "To improve operational resilience, RIC incorporates a rule-based fallback mechanism that guarantees uninterrupted triage during network partitions "
        "or external AI service outages. The fallback engine evaluates candidate commits using multi-domain keyword relevance, commit time proximity, "
        "and changed-file significance to produce reliable correlation intelligence without external dependencies. By synthesizing deterministic temporal "
        "filtering, automated code-diff analysis, high-speed AI inference, and fail-safe heuristic logic, the Release Incident Correlator delivers an "
        "indispensable diagnostic asset for modern DevOps and Site Reliability Engineering teams.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=24, line_spacing=1.5
    )

    table_syn = doc.add_table(rows=1, cols=2)
    table_syn.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_syn.autofit = False
    set_table_no_borders(table_syn)
    table_syn.cell(0, 0).width = Inches(3.1)
    table_syn.cell(0, 1).width = Inches(3.1)

    s0 = table_syn.cell(0, 0).paragraphs[0]
    run_s0 = s0.add_run("Name: Kavinraghul V")
    run_s0.font.bold = True
    run_s0.font.name = 'Times New Roman'

    s1 = table_syn.cell(0, 1).paragraphs[0]
    s1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_s1 = s1.add_run("Roll No: 23MSS020")
    run_s1.font.bold = True
    run_s1.font.name = 'Times New Roman'

    # =========================================================================
    # 6. TABLE OF CONTENTS (Exact Style & Borders Matching Sri Dharani & 23MSS011)
    # =========================================================================
    doc.add_page_break()
    add_p("TABLE OF CONTENTS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=18)

    toc_data = [
        ("S.NO", "CONTENTS", "PAGE NO"),
        ("1", "INTRODUCTION", ""),
        ("", "1.1 Problem Statement", "1"),
        ("", "1.2 Project Overview", "2"),
        ("", "1.3 Module Description", "3"),
        ("2", "SYSTEM ANALYSIS", ""),
        ("", "2.1 Existing System", "5"),
        ("", "2.2 Proposed System", "6"),
        ("3", "SYSTEM CONFIGURATION", ""),
        ("", "3.1 Hardware Specifications", "8"),
        ("", "3.2 Software Specifications", "8"),
        ("4", "SOFTWARE DESCRIPTION", ""),
        ("", "4.1 Frontend Technologies (React, Vite, Tailwind CSS, Ant Design)", "10"),
        ("", "4.2 Backend Framework (Python, FastAPI, Uvicorn, Pydantic)", "12"),
        ("", "4.3 Database Layer (PostgreSQL & SQLite Architecture)", "14"),
        ("", "4.4 AI & LLM Inference (Groq Cloud, Ollama, Heuristic Engine)", "15"),
        ("5", "SYSTEM DESIGN", ""),
        ("", "5.1 System Architecture & Flow Diagram", "17"),
        ("", "5.2 Data Flow Diagram (DFD Level 0 & Level 1)", "19"),
        ("", "5.3 Database Design / Table Design", "21"),
        ("", "5.4 Input Design", "24"),
        ("", "5.5 Output Design", "25"),
        ("6", "SYSTEM TESTING AND IMPLEMENTATION", ""),
        ("", "6.1 System Testing & Test Cases", "27"),
        ("", "6.2 System Implementation & Deployment", "30"),
        ("7", "SCOPE FOR FUTURE ENHANCEMENT", "32"),
        ("8", "CONCLUSION", "33"),
        ("", "BIBLIOGRAPHY", "34"),
        ("", "APPENDIX", "35"),
        ("", "A] Screenshots", "35"),
        ("", "B] Sample Coding", "44"),
    ]

    table_toc = doc.add_table(rows=len(toc_data), cols=3)
    table_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_toc.autofit = False
    set_table_borders(table_toc, color="000000", sz="4", val="single")

    col_widths = [Inches(0.6), Inches(4.8), Inches(1.0)]
    for r_idx, row in enumerate(table_toc.rows):
        is_header = (r_idx == 0)
        sno, title, pno = toc_data[r_idx]
        for c_idx, cell in enumerate(row.cells):
            cell.width = col_widths[c_idx]
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            if c_idx == 0 or c_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            run = p.add_run(toc_data[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            # Bold major chapter rows and table header
            if is_header or (sno and sno.isdigit()) or title in ["BIBLIOGRAPHY", "APPENDIX"]:
                run.font.bold = True

    # =========================================================================
    # SECTION 2: MAIN BODY (Chapters 1 to 8, Bibliography, Appendix)
    # Starts on a new section so page numbering restarts at 1!
    # =========================================================================
    sec_body = doc.add_section(WD_SECTION.NEW_PAGE)
    sec_body.top_margin = Inches(1.0)
    sec_body.bottom_margin = Inches(1.0)
    sec_body.left_margin = Inches(1.25)
    sec_body.right_margin = Inches(1.0)
    sec_body.page_width = Inches(8.5)
    sec_body.page_height = Inches(11.0)

    # Unlink footer from preliminary pages & restart page numbering at 1
    sec_body.footer.is_linked_to_previous = False
    sectPr = sec_body._sectPr
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:start'), '1')
    sectPr.append(pgNumType)

    # Add dynamic centered page number to body footer
    footer_p = sec_body.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run()
    footer_run.font.name = 'Times New Roman'
    footer_run.font.size = Pt(11)
    fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    footer_run._r.append(fldSimple)

    # ----------------------------------------------------
    # CHAPTER 1: INTRODUCTION
    # ----------------------------------------------------
    add_p("CHAPTER 1", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=16, space_after=4)
    add_p("INTRODUCTION", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=0, space_after=18)

    add_section_heading("1.1 PROBLEM STATEMENT")
    add_p(
        "In modern cloud-native software enterprises, software engineering teams execute continuous integration and continuous deployment "
        "(CI/CD) pipelines to release updates, feature additions, security patches, and performance optimizations on a daily or hourly basis. "
        "While rapid release cycles dramatically enhance business agility and responsiveness to user demand, they simultaneously amplify the risk "
        "of introducing unintended software regressions, subtle defects, and catastrophic production outages.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "When an incident strikes production—such as broken user login authentication, silent payment gateway drops, Elasticsearch index corruption, "
        "or degraded database connection pools—Site Reliability Engineering (SRE) and DevOps personnel are placed under immense stress to triage and resolve "
        "the outage. Traditionally, resolving an outage requires engineers to manually inspect application monitoring dashboards, scour log aggregators, "
        "open Git commit logs across multiple microservice repositories, and manually review hundreds of code lines in recent Git diffs.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "This manual troubleshooting workflow is fraught with substantial shortcomings: it is error-prone, cognitively draining, and excessively slow. "
        "In high-throughput enterprise systems, an outage lingering for hours can translate into tens of thousands of dollars in lost revenue, brand damage, "
        "and SLA penalties. Consequently, there is an urgent industry requirement for an automated, intelligent system that can instantly bridge the gap "
        "between reported production incidents and recent code deployments, pinpointing the exact Git commit responsible for the failure in seconds.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("1.2 PROJECT OVERVIEW")
    add_p(
        "Release Incident Correlator (RIC) is an advanced, AI-powered Site Reliability Engineering (SRE) and DevOps intelligence system designed "
        "to automatically identify the root-cause Git commit responsible for any reported production incident. RIC eliminates manual log hunting and "
        "diff reading by combining deterministic temporal filtering, code diff inspection, high-speed Large Language Model (LLM) inference, "
        "and a fail-safe heuristic fallback engine.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "The system operates by ingesting two primary data sources: incident logs formatted as CSV files (containing incident identifiers, titles, descriptions, "
        "ISO-8601 timestamps, and severity levels) and Git commit records formatted as JSON files (containing commit hashes, author identities, commit timestamps, "
        "commit messages, lists of changed files, and complete unified diff payloads). Upon receiving an analysis request, RIC normalizes all timestamps to UTC, "
        "prunes irrelevant historical commits using a customizable time-window filter (default: 6 hours before incident occurrence), and isolates the closest candidate deployments.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "To achieve near-instantaneous response times without incurring exorbitant API token expenses, RIC bundles candidate commits into a single, compact "
        "structured prompt dispatched to Groq's Language Processing Unit (LPU) cloud infrastructure running the state-of-the-art LLaMA 3.1 8B Instant model "
        "(or alternatively, a local Ollama server running LLaMA 3.1 on-premises). The model evaluates code semantic changes against the incident symptoms, "
        "returning a structured JSON payload detailing the suspect commit hash, a statistical confidence score (0-100%), blast radius impact rating (Low, Medium, "
        "High, Critical), a root-cause explanation, and concrete remediation/rollback instructions. The results are persisted in a relational database "
        "(PostgreSQL / SQLite 3) and presented through a reactive, glassmorphic dashboard built using React 18, Vite, Ant Design, and Tailwind CSS.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("1.3 MODULE DESCRIPTION")
    add_p(
        "The Release Incident Correlator architecture is partitioned into seven distinct, highly cohesive functional modules that operate collaboratively "
        "to deliver end-to-end incident triage:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("1. Incident Ingestion Module:")
    add_p(
        "Responsible for handling multipart file uploads of production incident tickets in CSV format. The module executes UTF-8 BOM sanitization, "
        "validates required header schemas (title, timestamp), handles optional fields (description, severity, external ID), and executes bulk batch insertions "
        "into the relational persistence layer while recording audit session metadata in upload_sessions.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("2. Commit & Diff Ingestion Module:")
    add_p(
        "Parses structured JSON payloads representing Git repository histories. Validates JSON schema arrays, ensures mandatory fields (commit_id, timestamp) "
        "are intact, serializes changed file paths into JSON string arrays, and stores unified diff text blocks securely into the database.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("3. Time-Window Correlation & Candidate Pruning Module:")
    add_p(
        "Performs temporal normalization on ISO-8601 timestamps and filters candidate commits strictly deployed within a configurable backward window "
        "(T_incident - window_hours <= T_commit <= T_incident). Sorts candidate commits in ascending order by deployment delta, ensuring that the most recent "
        "code changes deployed immediately prior to an incident receive prioritized evaluation.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("4. AI-Powered Diff Inspection Module:")
    add_p(
        "Constructs a token-lean, structured prompt aggregating the incident description and the top three candidate commit diffs (truncated to 400 characters per candidate). "
        "Executes non-blocking asynchronous inferences via the Groq SDK (LLaMA 3.1 8B Instant) or local Ollama HTTP endpoints. Employs a robust JSON extraction parser "
        "that strips markdown code fences and isolates balanced JSON objects containing confidence metrics, explanations, and remediation advice.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("5. Deterministic Heuristic Fallback Engine:")
    add_p(
        "Guarantees continuous 100% operational availability during network partitions or cloud LLM rate-limit events. Uses rule-based domain keyword matching "
        "(across login, payment, database, email, search, and API vocabularies), time-proximity bonus calculations, and suspicious code marker detection "
        "(todo, fixme, bug, warning, hack) to derive confidence scores and blast radius assessments without external dependencies.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("6. Interactive Visualizer & Report Generator Module:")
    add_p(
        "Frontend interface featuring animated circular confidence progress rings, pulsing blast radius badges, and an interactive syntax-highlighted Git diff modal. "
        "Provides one-click export mechanisms allowing engineers to download analysis reports as structured plain text (.txt) files or serialized JSON.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("7. Conversational AI Assistant Module (\"Sparky ✨\"):")
    add_p(
        "A floating, interactive chatbot embedded throughout the user interface. Powered by an SRE persona system prompt, Sparky provides instant answers "
        "to technical queries regarding data upload specifications, blast radius definitions, and AI correlation mechanics.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 2: SYSTEM ANALYSIS
    # ----------------------------------------------------
    add_chapter_heading("2", "SYSTEM ANALYSIS")

    add_section_heading("2.1 EXISTING SYSTEM")
    add_p(
        "In existing enterprise software maintenance workflows, incident triage is largely conducted manually by on-call engineers, developers, and DevOps staff. "
        "When an alerting platform (such as Datadog, New Relic, or Grafana) triggers an alarm indicating elevated 5xx error rates, service latency spikes, or failed background jobs, "
        "the triage team follows a fragmented and repetitive investigation procedure:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "1. Manual Dashboard Inspection: Engineers log into centralized monitoring systems to inspect metrics graphs and pinpoint the approximate time when errors began occurring.\n"
        "2. Manual Log Searching: Engineers log into log aggregation platforms (e.g., Splunk, ELK Stack, CloudWatch) and write ad-hoc queries searching for exception stack traces.\n"
        "3. Manual Git Log Browsing: Engineers access GitHub, GitLab, or terminal Git utilities to manually read recent commit titles and diffs deployed around the incident window.\n"
        "4. Subjective Guesswork & Communication Overhead: Developers debate over Slack or Microsoft Teams regarding which microservice or code change could have caused the regression.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("Disadvantages of the Existing System:")
    add_p(
        "• Prolonged Mean Time to Resolution (MTTR): Manual inspection across disconnected systems routinely requires 45 minutes to several hours, prolonging costly outages.\n"
        "• Extreme Cognitive Overload: Reviewing dozens of commits and thousands of lines of unified diffs under intense operational pressure frequently causes engineers to overlook subtle bugs.\n"
        "• High Financial Impact: Production downtime in enterprise e-commerce, banking, or SaaS platforms causes substantial revenue leakage and SLA contract breach penalties.\n"
        "• Lack of Standardized Rollback Advice: Manual triage rarely generates structured, documented remediation instructions, resulting in disorganized rollback attempts that risk worsening system instability.\n"
        "• Single Point of Failure on Tribal Knowledge: Troubleshooting success depends heavily on senior developers' personal familiarity with recent codebase changes.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("2.2 PROPOSED SYSTEM")
    add_p(
        "The proposed system, Release Incident Correlator (RIC), radically transforms incident troubleshooting by introducing an automated, intelligence-driven "
        "correlation platform. RIC unifies incident symptom reports and Git version control histories into an automated triage pipeline, applying mathematical temporal filtering "
        "and artificial intelligence to identify the culprit commit in under five seconds.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "Upon uploading incident tickets and commit diffs, RIC's backend pipeline immediately normalizes all timestamps and prunes the search space by evaluating only commits "
        "deployed within a customizable backward window (default: 6 hours before incident). The top candidate commits are bundled into a single-prompt inference call dispatched "
        "to high-throughput Large Language Models (LLaMA 3.1 8B Instant via Groq Cloud or local Ollama). If external AI connectivity is interrupted, RIC automatically falls back "
        "to a deterministic heuristic engine, evaluating domain keyword overlap, time proximity, and suspicious code markers to ensure zero downtime.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_subsection_heading("Advantages of the Proposed System:")
    add_p(
        "• Rapid Root-Cause Pinpointing: Reduces root-cause discovery time from hours to under 5 seconds, significantly lowering MTTR.\n"
        "• Intelligent Semantic Code Inspection: Large Language Models evaluate actual unified code diff semantics and logic changes, rather than merely relying on commit titles.\n"
        "• Quantified Confidence & Blast Radius: Provides transparent confidence percentages (0-100%) and categorizes blast radius risk (Low, Medium, High, Critical).\n"
        "• Actionable Remediation Guidance: Delivers concrete, copy-pasteable Git rollback commands and regression verification instructions.\n"
        "• Fail-Safe High Availability: Dual AI and rule-based heuristic fallback architecture guarantees 100% operational uptime.\n"
        "• Rich Interactive Visualizations: Modern glassmorphic interface with color-coded syntax diff viewing, animated progress bars, and historical auditing.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 3: SYSTEM CONFIGURATION
    # ----------------------------------------------------
    add_chapter_heading("3", "SYSTEM CONFIGURATION")

    add_section_heading("3.1 HARDWARE SPECIFICATIONS")
    add_p(
        "The hardware specifications describe the minimum and recommended computing resource requirements necessary to execute the Release Incident Correlator system efficiently:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    table_hw = doc.add_table(rows=6, cols=3)
    table_hw.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hw.autofit = False
    set_table_borders(table_hw, color="000000", sz="4", val="single")

    hw_data = [
        ("HARDWARE COMPONENT", "MINIMUM REQUIREMENT", "RECOMMENDED SPECIFICATION"),
        ("Processor (CPU)", "Intel Core i3 / AMD Ryzen 3 (2.0 GHz)", "Intel Core i7 / AMD Ryzen 7 / Apple Silicon"),
        ("System Memory (RAM)", "4 GB RAM", "8 GB – 16 GB DDR4/DDR5 RAM"),
        ("Storage (Disk Space)", "500 MB free hard disk space", "5 GB SSD (if caching local LLMs via Ollama)"),
        ("Display Monitor", "1366 x 768 Resolution", "1920 x 1080 Full HD IPS Display"),
        ("Network Connectivity", "Standard Broadband (512 Kbps)", "High-Speed Internet (for Groq Cloud API inference)"),
    ]
    for r_idx, row in enumerate(table_hw.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(2.1)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(hw_data[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            if is_header:
                run.font.bold = True

    add_section_heading("3.2 SOFTWARE SPECIFICATIONS")
    add_p(
        "The software specifications define the operating systems, programming languages, application frameworks, database servers, and runtime tools "
        "required to build, execute, and host the Release Incident Correlator platform:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    table_sw = doc.add_table(rows=8, cols=3)
    table_sw.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sw.autofit = False
    set_table_borders(table_sw, color="000000", sz="4", val="single")

    sw_data = [
        ("SOFTWARE LAYER", "TECHNOLOGY / TOOL", "VERSION / SPECIFICATION"),
        ("Operating System", "Microsoft Windows / Linux / macOS", "Windows 10/11 / Ubuntu 22.04 LTS / macOS Sonoma"),
        ("Frontend Runtime & Build", "Node.js & npm / Vite", "Node.js v18.x or v20.x, Vite v5.3.0"),
        ("Frontend UI Framework", "React.js / Tailwind CSS / Ant Design", "React v18.3.1, Tailwind v4.3.0, antd v5.18.0"),
        ("Backend Framework", "Python / FastAPI / Uvicorn", "Python 3.10+, FastAPI v0.111.0, Uvicorn v0.30.1"),
        ("Data Validation", "Pydantic", "Pydantic v2.7.3"),
        ("Database Management System", "PostgreSQL / SQLite 3", "PostgreSQL v15+ (psycopg2) / SQLite 3.x (ric.db)"),
        ("AI / LLM Inference Platforms", "Groq Cloud LPU SDK / Ollama", "Groq SDK v0.9.0+ (LLaMA 3.1 8B) / Ollama Local"),
    ]
    for r_idx, row in enumerate(table_sw.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(2.1)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(sw_data[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)
            if is_header:
                run.font.bold = True

    # ----------------------------------------------------
    # CHAPTER 4: SOFTWARE DESCRIPTION
    # ----------------------------------------------------
    add_chapter_heading("4", "SOFTWARE DESCRIPTION")

    add_section_heading("4.1 FRONTEND TECHNOLOGIES")
    add_p(
        "The frontend application is constructed as an ultra-modern, highly reactive Single Page Application (SPA) designed to deliver an intuitive user experience "
        "and instantaneous feedback during high-stress operational incidents.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("React.js (v18.3.1):")
    add_p(
        "React is an open-source, component-based JavaScript library maintained by Meta. In this project, React powers the modular UI architecture, managing reactive state "
        "across data upload, real-time progress bars, and analysis rendering. React hooks including useState, useEffect, useRef, and useContext are heavily utilized "
        "to manage asynchronous state transitions without page reloads.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Vite (v5.3.0):")
    add_p(
        "Vite is a next-generation frontend build tool that leverages native ES modules in the browser. It delivers instantaneous Hot Module Replacement (HMR) during development "
        "and outputs highly optimized production bundles. Vite's development server is configured with a custom proxy that routes /api requests to the FastAPI backend at "
        "http://localhost:8000 while enforcing a 600,000ms proxy timeout to accommodate slower local Ollama CPU runs without socket aborts.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Ant Design (antd v5.18.0) & Tailwind CSS (v4.3.0):")
    add_p(
        "Ant Design provides enterprise-grade UI components including Progress circular gauges, Dropdown menus, Modal dialogs, Tooltips, Collapse panels, and notification Messages. "
        "Tailwind CSS v4 complements Ant Design by supplying modern utility classes, glassmorphic dark-theme gradients, and custom CSS animations (sparkle text, floating particles).",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Axios (v1.7.2):")
    add_p(
        "Axios handles asynchronous HTTP communication with the backend. It is configured with response interceptors that normalize network errors and display actionable timeout alerts "
        "if an on-premises Ollama model takes extended time to generate predictions.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("4.2 BACKEND FRAMEWORK")
    add_p(
        "The backend is engineered using Python 3 and FastAPI, adhering to high-performance, asynchronous web architecture principles.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("FastAPI (v0.111.0):")
    add_p(
        "FastAPI is a modern, fast (high-performance) web framework for building APIs with Python based on standard Python type hints. It automatically produces interactive OpenAPI (Swagger) "
        "documentation and supports native asynchronous endpoints (async def), allowing high concurrency during long-running AI inference and database operations.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Uvicorn (v0.30.1):")
    add_p(
        "Uvicorn is a lightning-fast ASGI (Asynchronous Server Gateway Interface) web server implementation for Python. It serves the FastAPI application, utilizing uvloop and httptools "
        "to manage non-blocking socket connections and concurrent HTTP requests.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Pydantic (v2.7.3):")
    add_p(
        "Pydantic provides data validation and settings management using Python type annotations. It enforces strict request and response schemas (HealthResponse, UploadResponse, "
        "StatsResponse, AnalysisResult, RunSummary), ensuring complete type safety across all REST endpoints.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("4.3 DATABASE LAYER")
    add_p(
        "The persistence layer is abstracted in backend/database.py to provide transparent, dual-database compatibility:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "• PostgreSQL (Cloud Mode): Configured via the CONNECTION_STRING_PSQL environment variable and accessed through psycopg2 with RealDictCursor. Serves as the primary scalable "
        "database for production cloud deployments (such as Render or AWS RDS).\n"
        "• SQLite 3 (Local Fallback Mode): If no PostgreSQL connection string is supplied, the system automatically falls back to an embedded SQLite database stored at backend/ric.db. "
        "It applies PRAGMA foreign_keys = ON and configures a 30.0-second lock timeout to prevent concurrency contention during rapid local file ingestion.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("4.4 ARTIFICIAL INTELLIGENCE & LLM INFERENCE")
    add_p(
        "RIC leverages cutting-edge generative AI models and custom heuristic engineering to evaluate code diffs against incident symptoms:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Groq Cloud Language Processing Units (LPUs):")
    add_p(
        "Groq's LPU hardware architecture delivers deterministic, ultra-fast LLM token generation speeds. Using the official Groq Python SDK, RIC queries the LLaMA 3.1 8B Instant model "
        "with temperature=0.1 and max_tokens=350, completing full code-diff analysis in under 1 to 2 seconds.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Ollama Local LLM Integration:")
    add_p(
        "For enterprise organizations with strict data-privacy regulations where proprietary code cannot leave on-premises servers, RIC supports local Ollama endpoints (http://localhost:11434) "
        "running the llama3.1:8b model via asynchronous HTTP POST requests powered by httpx.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_subsection_heading("Deterministic Heuristic Fallback Engine:")
    add_p(
        "To ensure continuous operations when cloud APIs or local models are offline, RIC implements a rule-based domain keyword scoring engine. It analyzes term frequency across six domain dictionaries "
        "(login, payment, database, email, search, API), rewards deployment proximity (+15% if deployed <30m before incident), and penalizes suspicious code markers (todo, fixme, bug, hack) "
        "to deliver reliable correlation without external AI dependencies.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 5: SYSTEM DESIGN
    # ----------------------------------------------------
    add_chapter_heading("5", "SYSTEM DESIGN")

    add_section_heading("5.1 SYSTEM ARCHITECTURE & FLOW DIAGRAM")
    add_p(
        "The Release Incident Correlator follows a decoupled, four-tier microservice-inspired architecture comprising the Presentation Tier, Application Tier, "
        "Inference Tier, and Persistence Tier. Figure 5.1 illustrates the end-to-end component architecture and runtime data flow:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    if os.path.exists("images/architecture.png"):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture("images/architecture.png", width=Inches(5.6))
        add_p("Figure 5.1: Overall System Architecture of Release Incident Correlator", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True, size=10, space_before=2, space_after=14)

    add_p(
        "Figure 5.2 demonstrates the execution workflow from client file upload, time-window candidate filtering, Groq/Ollama LLM inference, and fallback decision paths:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    if os.path.exists("images/working.png"):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run_img = p_img.add_run()
        run_img.add_picture("images/working.png", width=Inches(5.6))
        add_p("Figure 5.2: End-to-End Processing and Triage Workflow", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True, size=10, space_before=2, space_after=14)

    add_section_heading("5.2 DATA FLOW DIAGRAMS")
    add_subsection_heading("DFD Level 0 (Context Level Diagram):")
    add_p(
        "At Level 0, the external entity (Site Reliability Engineer / DevOps Developer) interacts with the centralized Release Incident Correlator system. "
        "The user submits incident tickets (CSV) and Git commit histories (JSON). The system processes the correlation and outputs confidence-ranked root-cause analysis "
        "reports and rollback recommendations.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    p_dfd0 = add_p(
        "               ┌────────────────────────┐\n"
        "               │      SRE / DEVOPS      │\n"
        "               │       ENGINEER         │\n"
        "               └───────────┬────────────┘\n"
        "         Incidents CSV /   │    ▲  Root-Cause Reports,\n"
        "         Commits JSON      │    │  Rollback Recommendations,\n"
        "                           ▼    │  Diff Visualizations\n"
        "               ┌────────────────────────┐\n"
        "               │          0.0           │\n"
        "               │    RELEASE INCIDENT    │\n"
        "               │       CORRELATOR       │\n"
        "               └────────────────────────┘",
        align=WD_ALIGN_PARAGRAPH.CENTER, size=10, space_before=6, space_after=14, line_spacing=1.0
    )
    p_dfd0.runs[0].font.name = 'Courier New'

    add_subsection_heading("DFD Level 1 (Detailed Functional Process Level):")
    add_p(
        "Level 1 decomposes the RIC system into four core functional processes: Data Ingestion & Schema Validation (1.0), Time-Window Candidate Pruning (2.0), "
        "AI / Heuristic Root Cause Inference (3.0), and Analysis Visualization & Exporting (4.0).",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    p_dfd1 = add_p(
        " [User Upload] ──► (1.0 Ingestion & Validation) ──► [DB: incidents / commits]\n"
        "                                                             │\n"
        "                                                             ▼\n"
        "                     [DB Batch Query] ──► (2.0 Time-Window Pruning: ±6h)\n"
        "                                                             │\n"
        "                                                             ▼\n"
        "                                            [Top 3 Candidate Commits]\n"
        "                                                             │\n"
        "                                                             ▼\n"
        "                                            (3.0 AI / Heuristic Inference)\n"
        "                                             ├── Groq Cloud LPU\n"
        "                                             ├── Ollama Local Model\n"
        "                                             └── Rule-Based Fallback\n"
        "                                                             │\n"
        "                                                             ▼\n"
        " [Export TXT/JSON] ◄── (4.0 UI Visualization) ◄── [DB: analyses Table]",
        align=WD_ALIGN_PARAGRAPH.CENTER, size=9, space_before=6, space_after=14, line_spacing=1.0
    )
    p_dfd1.runs[0].font.name = 'Courier New'

    add_section_heading("5.3 DATABASE DESIGN / TABLE DESIGN")
    add_p(
        "The relational schema contains four interrelated tables designed with primary keys, cascading foreign keys, and indexes for high-speed retrieval:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # Table 1: incidents
    add_subsection_heading("Table 5.3.1: incidents Table Schema")
    t_inc = doc.add_table(rows=8, cols=5)
    t_inc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_inc.autofit = False
    set_table_borders(t_inc, color="000000", sz="4", val="single")

    inc_cols = [
        ("COLUMN NAME", "DATA TYPE", "NULLABLE", "KEY / CONSTRAINT", "DESCRIPTION"),
        ("id", "BIGSERIAL / INT", "No", "PRIMARY KEY", "Auto-increment unique incident ID"),
        ("incident_ref", "TEXT", "No", "INDEX (idx_inc_ref)", "External incident ticket reference code"),
        ("title", "TEXT", "No", "NOT NULL", "Brief summary of the production outage"),
        ("description", "TEXT", "No", "DEFAULT ''", "Detailed symptom description and error logs"),
        ("timestamp", "TEXT", "No", "NOT NULL", "ISO-8601 UTC timestamp of incident onset"),
        ("severity", "TEXT", "No", "DEFAULT 'medium'", "Severity tier: low, medium, high, critical"),
        ("uploaded_at", "TIMESTAMPTZ / DATETIME", "No", "DEFAULT CURRENT_TIMESTAMP", "Batch audit ingestion timestamp"),
    ]
    for r_idx, row in enumerate(t_inc.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(1.25)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(inc_cols[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9.5)
            if is_header:
                run.font.bold = True

    # Table 2: commits
    add_p("", space_before=6, space_after=6)
    add_subsection_heading("Table 5.3.2: commits Table Schema")
    t_com = doc.add_table(rows=10, cols=5)
    t_com.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_com.autofit = False
    set_table_borders(t_com, color="000000", sz="4", val="single")

    com_cols = [
        ("COLUMN NAME", "DATA TYPE", "NULLABLE", "KEY / CONSTRAINT", "DESCRIPTION"),
        ("id", "BIGSERIAL / INT", "No", "PRIMARY KEY", "Auto-increment unique commit ID"),
        ("commit_hash", "TEXT", "No", "INDEX (idx_com_hash)", "Git SHA commit hash (short or full)"),
        ("author", "TEXT", "No", "DEFAULT 'Unknown'", "Author name of the Git commit"),
        ("email", "TEXT", "No", "DEFAULT ''", "Author email address"),
        ("message", "TEXT", "No", "DEFAULT ''", "Git commit message header"),
        ("timestamp", "TEXT", "No", "NOT NULL", "ISO-8601 UTC deployment timestamp"),
        ("branch", "TEXT", "No", "DEFAULT 'main'", "Target deployment Git branch"),
        ("files_changed", "TEXT", "No", "DEFAULT '[]'", "JSON serialized array of modified file paths"),
        ("diff", "TEXT", "No", "DEFAULT ''", "Unified diff text representing code changes"),
    ]
    for r_idx, row in enumerate(t_com.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(1.25)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(com_cols[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9.5)
            if is_header:
                run.font.bold = True

    # Table 3: analyses
    add_p("", space_before=6, space_after=6)
    add_subsection_heading("Table 5.3.3: analyses Table Schema")
    t_ana = doc.add_table(rows=11, cols=5)
    t_ana.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_ana.autofit = False
    set_table_borders(t_ana, color="000000", sz="4", val="single")

    ana_cols = [
        ("COLUMN NAME", "DATA TYPE", "NULLABLE", "KEY / CONSTRAINT", "DESCRIPTION"),
        ("id", "BIGSERIAL / INT", "No", "PRIMARY KEY", "Unique analysis correlation record ID"),
        ("run_id", "TEXT", "No", "INDEX (idx_ana_run)", "UUID v4 grouping multi-incident batch runs"),
        ("run_label", "TEXT", "Yes", "DEFAULT NULL", "Formatted human-readable UTC timestamp label"),
        ("incident_id", "BIGINT / INT", "Yes", "FK (incidents.id) CASCADE", "Referenced incident foreign key"),
        ("commit_id", "BIGINT / INT", "Yes", "FK (commits.id) CASCADE", "Referenced root-cause commit foreign key"),
        ("author", "TEXT", "Yes", "DEFAULT NULL", "Author of the identified culprit commit"),
        ("confidence", "FLOAT / REAL", "No", "DEFAULT 0.0", "AI/Heuristic confidence score (0-100)"),
        ("explanation", "TEXT", "Yes", "DEFAULT NULL", "Root cause explanation explaining the failure"),
        ("recommendation", "TEXT", "Yes", "DEFAULT NULL", "Actionable remediation & rollback instructions"),
        ("blast_radius", "TEXT", "No", "DEFAULT 'medium'", "Blast radius: low, medium, high, critical"),
    ]
    for r_idx, row in enumerate(t_ana.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(1.25)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(ana_cols[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9.5)
            if is_header:
                run.font.bold = True

    add_section_heading("5.4 INPUT DESIGN")
    add_p(
        "Input design governs how external data enters the Release Incident Correlator system. RIC supports two structured input formats:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "1. Incidents CSV Input: Ingested through POST /upload/incidents. The CSV requires title and timestamp columns, with optional description, severity, and id columns.\n"
        "   Example Record: User Authentication Failure, JWT verify() throws 500 error on sync mode, 2024-03-12T10:45:00Z, critical\n\n"
        "2. Commits JSON Input: Ingested through POST /upload/commits. The JSON payload consists of an array of commit objects containing commit_id, timestamp, author, message, "
        "files_changed (array of file path strings), and diff (unified diff format).\n"
        "   Example Record: {\"commit_id\": \"a7d4e91\", \"author\": \"Sarah Chen\", \"message\": \"refactor: migrate JWT validation\", \"timestamp\": \"2024-03-12T09:58:00Z\", ...}",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    add_section_heading("5.5 OUTPUT DESIGN")
    add_p(
        "Output design dictates how analytical findings are presented to the engineering team. RIC delivers outputs across multiple visual and downloadable modalities:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "• Circular Confidence Ring: An interactive Ant Design progress ring color-coded based on confidence threshold (>=80% vibrant green, >=60% warning gold, <60% danger red).\n"
        "• Blast Radius Badge: A pulsing status pill classifying the blast radius as Low, Medium, High, or Critical.\n"
        "• Root Cause & Recommendation Cards: Two side-by-side analytical cards providing clear natural-language justifications and concrete Git rollback commands.\n"
        "• Syntax-Highlighted Diff Modal: An interactive modal displaying unified diff line numbers with green background (+ additions) and red background (- deletions).\n"
        "• Text Report & JSON Export: One-click export downloading the analysis summary formatted for incident post-mortems or downstream integration.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 6: SYSTEM TESTING & IMPLEMENTATION
    # ----------------------------------------------------
    add_chapter_heading("6", "SYSTEM TESTING AND IMPLEMENTATION")

    add_section_heading("6.1 SYSTEM TESTING & TEST CASES")
    add_p(
        "System testing was conducted across all modules to ensure functional correctness, data validation accuracy, time-window calculation precision, "
        "and seamless failover between AI inference and deterministic heuristic processing. Table 6.1 documents the comprehensive test suite executed for RIC:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    table_test = doc.add_table(rows=9, cols=5)
    table_test.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_test.autofit = False
    set_table_borders(table_test, color="000000", sz="4", val="single")

    test_cases = [
        ("TC ID", "TEST SCENARIO", "INPUT DATA", "EXPECTED RESULT", "STATUS"),
        ("TC-01", "Valid CSV Upload", "incidents.csv with title, timestamp", "HTTP 200, records inserted into DB", "PASSED"),
        ("TC-02", "Invalid CSV Upload", "CSV missing 'timestamp' header", "HTTP 400 'missing required columns'", "PASSED"),
        ("TC-03", "Valid JSON Upload", "commits.json with commit_id, diff", "HTTP 200, commits saved with diffs", "PASSED"),
        ("TC-04", "Time-Window Pruning", "Commits deployed 2h, 5h, and 12h prior", "Only 2h and 5h commits selected (<=6h)", "PASSED"),
        ("TC-05", "Groq AI Inference", "Incident symptom + 3 candidate diffs", "Model identifies root commit with >85% conf", "PASSED"),
        ("TC-06", "Heuristic Fallback Trigger", "Groq API key disabled / network drop", "Heuristic engine returns domain score & fix", "PASSED"),
        ("TC-07", "Interactive Diff Viewer", "Click 'See Diff' on suspect commit", "DiffModal displays line additions and deletions", "PASSED"),
        ("TC-08", "Run Deletion Cascade", "Delete run_id from History.jsx", "Cascading SQL delete removes all run records", "PASSED"),
    ]
    for r_idx, row in enumerate(table_test.rows):
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(1.25)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            if is_header:
                set_cell_shading(cell, "F2F2F2")
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(test_cases[r_idx][c_idx])
            run.font.name = 'Times New Roman'
            run.font.size = Pt(9.5)
            if is_header:
                run.font.bold = True
            elif c_idx == 4:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 130, 0)

    add_section_heading("6.2 SYSTEM IMPLEMENTATION & DEPLOYMENT")
    add_p(
        "The Release Incident Correlator is implemented as a cloud-ready, container-friendly web application:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "• Frontend Deployment: Hosted on Vercel (https://track-incident.vercel.app/). Configured via vercel.json with SPA rewrites that route all non-static paths to index.html.\n"
        "• Backend Deployment: Hosted on Render as a Python Web Service running Uvicorn (uvicorn main:app --host 0.0.0.0 --port 8000).\n"
        "• Local Execution Setup:\n"
        "  1. Backend: python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt && python main.py\n"
        "  2. Frontend: npm install && npm run dev (accessible at http://localhost:5173)\n"
        "• Environment Variables: Managed centrally via backend/.env (CONNECTION_STRING_PSQL, AI_PROVIDER=groq, GROQ_API_KEY, GROQ_MODEL, TIME_WINDOW_HOURS=6).",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 7: SCOPE FOR FUTURE ENHANCEMENT
    # ----------------------------------------------------
    add_chapter_heading("7", "SCOPE FOR FUTURE ENHANCEMENT")
    add_p(
        "While the current release of RIC successfully solves the core incident-commit correlation problem, several advanced technical enhancements "
        "are planned for future development iterations:",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "1. Real-Time Git Webhook Ingestion: Integrating GitHub and GitLab webhook listeners to ingest Git commits and pull request merges continuously in real time, "
        "eliminating manual commits.json generation.\n\n"
        "2. Automated CI/CD Rollback Automation: Providing an authorized 'Trigger Automated Rollback' action button that interacts with GitHub Actions or ArgoCD APIs "
        "to revert suspect commits and redeploy staging/production environments automatically.\n\n"
        "3. Distributed Multi-Repository Microservices Dependency Mapping: Constructing dependency graph trees across microservice repositories (e.g., frontend, auth service, "
        "payment queue) to trace cascading failures introduced across service boundaries.\n\n"
        "4. APM & Alerting Platform Webhook Listeners: Connecting directly with PagerDuty, Datadog, Sentry, and AWS CloudWatch webhooks to automatically initiate correlation "
        "the exact second an incident alert fires.\n\n"
        "5. Role-Based Access Control (RBAC) & Enterprise SSO: Implementing OAuth2 and SAML Single Sign-On (SSO) with granular permissions for SRE Leads, Developers, and Auditors.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # CHAPTER 8: CONCLUSION
    # ----------------------------------------------------
    add_chapter_heading("8", "CONCLUSION")
    add_p(
        "The Release Incident Correlator (RIC) project successfully addresses one of the most persistent, expensive, and stressful challenges in modern software engineering: "
        "pinpointing which code change caused a production incident. By replacing manual log hunting and subjective code review with an automated, intelligence-driven pipeline, "
        "RIC reduces Mean Time to Resolution (MTTR) from hours to under five seconds.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "Through a decoupled architecture combining React 18, Vite, Ant Design, Tailwind CSS, FastAPI, and PostgreSQL/SQLite, the platform delivers high throughput, "
        "aesthetic excellence, and operational resilience. The innovative integration of single-prompt candidate batching via Groq Cloud's LPU infrastructure optimizes token consumption "
        "and inference latency, while the deterministic heuristic engine provides a robust fail-safe mechanism guaranteeing uninterrupted triage availability.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )
    add_p(
        "In conclusion, RIC demonstrates the practical, transformative potential of combining modern software engineering standards with domain-specific artificial intelligence. "
        "The project stands as a complete, fully verified, and valuable addition to the modern Site Reliability Engineering and DevOps operational toolkit.",
        align=WD_ALIGN_PARAGRAPH.JUSTIFY
    )

    # ----------------------------------------------------
    # BIBLIOGRAPHY
    # ----------------------------------------------------
    doc.add_page_break()
    add_p("BIBLIOGRAPHY", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=18)

    references = [
        "1. Beyer, B., Jones, C., Petoff, J., & Murphy, N. R. (2016). Site Reliability Engineering: How Google Runs Production Systems. O'Reilly Media, Inc.",
        "2. Ramirez, S. (2024). FastAPI: Modern, High-Performance Web APIs in Python. Official Documentation. https://fastapi.tiangolo.com/",
        "3. Meta AI. (2024). The LLaMA 3 Herd of Models. Meta AI Research Technical Report. arXiv:2407.21783.",
        "4. Groq Inc. (2024). Language Processing Units (LPUs) Architecture and Cloud Inference APIs. https://groq.com/",
        "5. Chacon, S., & Straub, B. (2014). Pro Git (2nd ed.). Apress.",
        "6. React Core Team. (2024). React 18 Documentation. Meta Open Source. https://react.dev/",
        "7. Ant Design Team. (2024). An Enterprise-Class UI Design Language and React Component Library. https://ant.design/",
        "8. Wieruch, R. (2022). The Road to React: Your journey to master plain yet pragmatic React. Leanpub.",
        "9. Ollama Team. (2024). Get up and running with large language models locally. https://ollama.com/",
        "10. Python Software Foundation. (2024). Python Database API Specification v2.0 (PEP 249). https://peps.python.org/pep-0249/",
    ]
    for ref in references:
        add_p(ref, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=4, space_after=8, line_spacing=1.25)

    # ----------------------------------------------------
    # APPENDIX: A] SCREENSHOTS
    # ----------------------------------------------------
    doc.add_page_break()
    add_p("APPENDIX", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=8)
    add_p("A] SCREENSHOTS", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=18)

    screenshots = [
        ("images/Landing page.png", "Figure A.1: Interactive Landing Page with Hero Section and Animated Terminal"),
        ("images/dashboard.png", "Figure A.2: Real-Time Incident Correlator Dashboard and Upload Dropzones"),
        ("images/result.png", "Figure A.3: AI Analysis Results with Confidence Gauge and Blast Radius"),
        ("images/History.png", "Figure A.4: Historical Analysis Runs Management and Drilldown View"),
        ("images/Ai interface.png", "Figure A.5: Interactive AI Conversational Assistant ('Sparky ✨') Interface"),
        ("images/enterprises.png", "Figure A.6: Enterprise Architecture and Target Audience Showcase Section"),
        ("images/tables.png", "Figure A.7: Database Schema and Tables Overview"),
        ("images/incident db.png", "Figure A.8: Production Incident Records in Database"),
        ("images/incident 2 db.png", "Figure A.9: Detailed Incident Attributes and Timestamp Mapping"),
        ("images/analysis db.png", "Figure A.10: Stored AI Analysis and Confidence Scores in Relational Database"),
    ]

    for img_path, caption in screenshots:
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=Inches(5.6))
            add_p(caption, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True, size=10, space_before=2, space_after=16)

    # ----------------------------------------------------
    # APPENDIX: B] SAMPLE CODING
    # ----------------------------------------------------
    doc.add_page_break()
    add_p("APPENDIX", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=14, space_before=18, space_after=8)
    add_p("B] SAMPLE CODING", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=13, space_before=0, space_after=18)

    sample_code_files = [
        ("backend/services/correlator.py", "B.1: Candidate Correlation & Deterministic Heuristic Engine (correlator.py)", 1, 90),
        ("backend/services/ai_service.py", "B.2: AI Inference & Prompt Engineering Service (ai_service.py)", 1, 95),
        ("backend/routers/analysis.py", "B.3: Analysis Router & Pipeline Execution (analysis.py)", 1, 110),
    ]

    for file_path, heading, start_l, end_l in sample_code_files:
        if os.path.exists(file_path):
            add_subsection_heading(heading)
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()[start_l-1:end_l]
            code_text = "".join(lines)
            
            p_code = doc.add_paragraph()
            p_code.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_code.paragraph_format.space_before = Pt(4)
            p_code.paragraph_format.space_after = Pt(14)
            p_code.paragraph_format.line_spacing = 1.0
            
            run_c = p_code.add_run(code_text)
            run_c.font.name = 'Courier New'
            run_c.font.size = Pt(8.5)
            run_c.font.color.rgb = RGBColor(20, 20, 20)

    # Save to documentation_details folder and copy to root
    out_path_details = "documentation_details/RELEASE_INCIDENT_CORRELATOR_MAJOR_PROJECT_DOCUMENTATION.docx"
    out_path_root = "RELEASE_INCIDENT_CORRELATOR_MAJOR_PROJECT_DOCUMENTATION.docx"
    doc.save(out_path_details)
    doc.save(out_path_root)
    print(f"Successfully generated complete documentation at:\n  - {out_path_details}\n  - {out_path_root}")

if __name__ == '__main__':
    create_document()
