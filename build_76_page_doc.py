import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        r'<w:tcBorders {}><w:top w:val="{}" w:sz="{}" w:space="0" w:color="{}"/><w:bottom w:val="{}" w:sz="{}" w:space="0" w:color="{}"/><w:left w:val="{}" w:sz="{}" w:space="0" w:color="{}"/><w:right w:val="{}" w:sz="{}" w:space="0" w:color="{}"/></w:tcBorders>'.format(
            nsdecls('w'),
            kwargs.get('top_val', 'single'), kwargs.get('top_sz', '4'), kwargs.get('top_color', 'B0B0B0'),
            kwargs.get('bottom_val', 'single'), kwargs.get('bottom_sz', '4'), kwargs.get('bottom_color', 'B0B0B0'),
            kwargs.get('left_val', 'single'), kwargs.get('left_sz', '4'), kwargs.get('left_color', 'B0B0B0'),
            kwargs.get('right_val', 'single'), kwargs.get('right_sz', '4'), kwargs.get('right_color', 'B0B0B0')
        )
    )
    tcPr.append(tcBorders)

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def create_76_page_report():
    doc = docx.Document()

    # Set standard 1.0 inch margins all around
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Base Normal Style
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Times New Roman'
    normal_font.size = Pt(12)
    normal_font.color.rgb = RGBColor(0, 0, 0)

    # Helper functions
    def add_divider_page(title_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(260)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title_text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(18)
        r.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)

    def add_page_title(text, space_before=16, space_after=12):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_subheading(text, space_before=10, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_p(text, bold_prefix=None, space_after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Times New Roman'
            rb.font.size = Pt(12)
            rb.bold = True
            rb.font.color.rgb = RGBColor(0, 0, 0)
        rt = p.add_run(text)
        rt.font.name = 'Times New Roman'
        rt.font.size = Pt(12)
        rt.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_bullet(text, bold_prefix=None, space_after=3):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix)
            rb.font.name = 'Times New Roman'
            rb.font.size = Pt(12)
            rb.bold = True
            rb.font.color.rgb = RGBColor(0, 0, 0)
        rt = p.add_run(text)
        rt.font.name = 'Times New Roman'
        rt.font.size = Pt(12)
        rt.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_code_page(title, code_text):
        p_t = doc.add_paragraph()
        p_t.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_t.paragraph_format.space_before = Pt(14)
        p_t.paragraph_format.space_after = Pt(10)
        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(13)
        r_t.bold = True

        p_c = doc.add_paragraph()
        p_c.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_c.paragraph_format.space_before = Pt(0)
        p_c.paragraph_format.space_after = Pt(6)
        p_c.paragraph_format.line_spacing = 1.05
        r_c = p_c.add_run(code_text)
        r_c.font.name = 'Courier New'
        r_c.font.size = Pt(9.0)
        r_c.font.color.rgb = RGBColor(20, 20, 20)

    def add_table_data(headers, rows, col_widths=None):
        tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = tbl.rows[0].cells
        for idx, h_text in enumerate(headers):
            cell = hdr_cells[idx]
            cell.text = h_text
            set_cell_shading(cell, "EAEAEA")
            set_cell_border(cell)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                for run in p.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(10.5)
                    run.bold = True

        for r_idx, row_data in enumerate(rows):
            row_cells = tbl.rows[r_idx + 1].cells
            bg_color = "F9F9F9" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_data):
                cell = row_cells[c_idx]
                cell.text = str(val)
                set_cell_shading(cell, bg_color)
                set_cell_border(cell)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(9.5)

        if col_widths:
            for row in tbl.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)

        sp_p = doc.add_paragraph()
        sp_p.paragraph_format.space_before = Pt(0)
        sp_p.paragraph_format.space_after = Pt(4)

    # =========================================================================
    # PAGE 1: TITLE PAGE / COVER
    # =========================================================================
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(36)
    p1.paragraph_format.space_after = Pt(12)
    r1 = p1.add_run("Online Examination System using Python")
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(18)
    r1.bold = True

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(12)
    p2.paragraph_format.space_after = Pt(6)
    r2 = p2.add_run("A major project work submitted in partial\nfulfillment of the requirements for the degree of")
    r2.font.name = 'Times New Roman'
    r2.font.size = Pt(12)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(8)
    p3.paragraph_format.space_after = Pt(8)
    r3 = p3.add_run("BACHELOR OF SCIENCE IN COMPUTER SCIENCE\n(3-B.Sc Computer Science)")
    r3.font.name = 'Times New Roman'
    r3.font.size = Pt(13)
    r3.bold = True

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(8)
    p4.paragraph_format.space_after = Pt(16)
    r4 = p4.add_run("to the\nPeriyar University, Salem – 636011\n\nBy")
    r4.font.name = 'Times New Roman'
    r4.font.size = Pt(12)

    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(0)
    p5.paragraph_format.space_after = Pt(20)
    r5 = p5.add_run("M.LOGESHWARI\n(Degree: 3-B.Sc Computer Science)")
    r5.font.name = 'Times New Roman'
    r5.font.size = Pt(14)
    r5.bold = True

    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(8)
    p6.paragraph_format.space_after = Pt(20)
    r6 = p6.add_run("Under the Guidance of\nMs. C. MOHANAPRIYA M.Sc.,\nHead of Department\nDepartment of Computer Science\n\nPACHAMUTHU COLLEGE OF ARTS AND SCIENCE FOR WOMEN\n(AFFILIATED TO PERIYAR UNIVERSITY)\nDHARMAPURI – 636701\nSEPTEMBER – 2026")
    r6.font.name = 'Times New Roman'
    r6.font.size = Pt(12)

    # =========================================================================
    # PAGE 2: DIVIDER - CERTIFICATE
    # =========================================================================
    doc.add_page_break()
    add_divider_page("CERTIFICATE")

    # =========================================================================
    # PAGE 3: BONAFIDE CERTIFICATE DETAILS
    # =========================================================================
    doc.add_page_break()
    p_c_hdr = doc.add_paragraph()
    p_c_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c_hdr.paragraph_format.space_before = Pt(16)
    p_c_hdr.paragraph_format.space_after = Pt(14)
    r_ch = p_c_hdr.add_run("PACHAMUTHU COLLEGE OF ARTS AND SCIENCE FOR WOMEN\n(AFFILIATED TO PERIYAR UNIVERSITY)\nDHARMAPURI – 636701\n\nPROJECT WORK\nSEPTEMBER – 2026\n\nOnline Examination System using Python\n\nBonafide work done\nBY\nM.LOGESHWARI\n(Degree: 3-B.Sc Computer Science)")
    r_ch.font.name = 'Times New Roman'
    r_ch.font.size = Pt(12)
    r_ch.bold = True

    add_p("A Project Work Submitted In Partial Fulfillment Of the Requirement For the Degree Of BACHELOR OF SCIENCE IN COMPUTER SCIENCE To The Periyar University, Salem-11.", align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("This is to certify that the work described in this project report entitled \"Online Examination System using Python\" is the bonafide work done by M.LOGESHWARI under my supervision and guidance during the academic year 2025–2026.")

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_sig.paragraph_format.space_before = Pt(36)
    p_sig.paragraph_format.space_after = Pt(16)
    r_sig = p_sig.add_run("INTERNAL GUIDE                                                HEAD OF THE DEPARTMENT\n(Ms. C. MOHANAPRIYA, M.Sc.)                           (Department of Computer Science)\n\n\nSUBMITTED FOR THE VIVA-VOCE EXAMINATION HELD ON: ____________________\n\n\nINTERNAL EXAMINER                                            EXTERNAL EXAMINER")
    r_sig.font.name = 'Times New Roman'
    r_sig.font.size = Pt(11)
    r_sig.bold = True

    # =========================================================================
    # PAGE 4: DIVIDER - ACKNOWLEDGMENT
    # =========================================================================
    doc.add_page_break()
    add_divider_page("ACKNOWLEDGMENT")

    # =========================================================================
    # PAGE 5: ACKNOWLEDGMENT CONTENT
    # =========================================================================
    doc.add_page_break()
    p_ack_hdr = doc.add_paragraph()
    p_ack_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ack_hdr.paragraph_format.space_before = Pt(20)
    p_ack_hdr.paragraph_format.space_after = Pt(18)
    r_ah = p_ack_hdr.add_run("ACKNOWLEDGEMENT")
    r_ah.font.name = 'Times New Roman'
    r_ah.font.size = Pt(14)
    r_ah.bold = True

    add_p("At the outset, I offer my humble prayers to the Almighty and to my parents for giving me the strength, perseverance, and determination not only in the pursuit of this task but also in all walks of life.")
    add_p("I express my sincere thanks to our honourable Principal Dr. J. THAVAMANI, M.COM., M.Phil., Ph.D. for providing the necessary facilities, academic resources, and encouraging atmosphere to carry out my project work successfully.")
    add_p("My wholehearted and profound gratitude to Ms. C. MOHANAPRIYA, M.Sc., Head of the Department, for her cooperation, guidance, and continuous motivation that she has given me at every step throughout this project.")
    add_p("My wholehearted and proud gratitude goes to my project guide Ms. C. MOHANAPRIYA, M.Sc., for supervising me in choosing this project topic, structuring system modules, and enlightening my endeavour throughout this project work.")
    add_p("I thank all the staff members of the Department of Computer Science for their valuable guidance, academic support, and constructive feedback. Finally, my gratitude goes to my parents, friends, and everyone who has made direct and indirect contributions to make this project a grand success.")

    # =========================================================================
    # PAGE 6: DIVIDER - CONTENTS
    # =========================================================================
    doc.add_page_break()
    add_divider_page("CONTENTS")

    # =========================================================================
    # PAGE 7: INDEX / CONTENTS TABLE PART 1
    # =========================================================================
    doc.add_page_break()
    p_idx_hdr = doc.add_paragraph()
    p_idx_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_idx_hdr.paragraph_format.space_before = Pt(16)
    p_idx_hdr.paragraph_format.space_after = Pt(14)
    r_ih = p_idx_hdr.add_run("CONTENTS")
    r_ih.font.name = 'Times New Roman'
    r_ih.font.size = Pt(14)
    r_ih.bold = True

    idx_headers_1 = ["CHAPTER NO", "TITLE", "PAGE NO"]
    idx_rows_1 = [
        ["", "SYNOPSIS", "11"],
        ["1", "INTRODUCTION", "13"],
        ["", "  1.1 SYSTEM SPECIFICATION", "14"],
        ["", "    1.1.1 HARDWARE CONFIGURATION", "14"],
        ["", "    1.1.2 SOFTWARE CONFIGURATION", "15"],
        ["2", "SYSTEM STUDY", "19"],
        ["", "  2.1 EXISTING SYSTEM", "20"],
        ["", "    2.1.1 DESCRIPTION", "20"],
        ["", "    2.1.2 DRAWBACKS", "20"],
        ["", "  2.2 PROPOSED SYSTEM", "21"],
        ["", "    2.2.1 DESCRIPTION", "21"],
        ["", "    2.2.2 FEATURES", "23"],
        ["", "  ADVANTAGES OF THE PROPOSED SYSTEM", "24"]
    ]
    add_table_data(idx_headers_1, idx_rows_1, col_widths=[1.5, 4.0, 1.0])

    # =========================================================================
    # PAGE 8: INDEX / CONTENTS TABLE PART 2
    # =========================================================================
    doc.add_page_break()
    p_idx_hdr2 = doc.add_paragraph()
    p_idx_hdr2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_idx_hdr2.paragraph_format.space_before = Pt(16)
    p_idx_hdr2.paragraph_format.space_after = Pt(14)
    r_ih2 = p_idx_hdr2.add_run("CONTENTS (Continued)")
    r_ih2.font.name = 'Times New Roman'
    r_ih2.font.size = Pt(14)
    r_ih2.bold = True

    idx_headers_2 = ["CHAPTER NO", "TITLE", "PAGE NO"]
    idx_rows_2 = [
        ["3", "SYSTEM DESIGN AND DEVELOPMENT", "26"],
        ["", "  3.1 FILE DESIGN", "26"],
        ["", "  3.2 INPUT DESIGN", "30"],
        ["", "  3.3 OUTPUT DESIGN", "32"],
        ["", "  3.4 DATABASE DESIGN", "35"],
        ["", "  3.5 SYSTEM DEVELOPMENT", "37"],
        ["", "  3.6 DESCRIPTION OF MODULES", "38"],
        ["4", "TESTING AND IMPLEMENTATION", "42"],
        ["", "  4.1 TEST CASES EXAMPLE", "42"],
        ["", "  4.2 IMPLEMENTATION", "48"],
        ["", "  4.3 DEPLOYMENT OPTIONS", "50"],
        ["5", "CONCLUSION", "73"],
        ["6", "BIBLIOGRAPHY", "76"],
        ["7", "APPENDICES", "51"],
        ["", "  A) SAMPLE CODING", "51"],
        ["", "  B) SAMPLE INPUT", "63"],
        ["", "  C) SAMPLE OUTPUT", "66"]
    ]
    add_table_data(idx_headers_2, idx_rows_2, col_widths=[1.5, 4.0, 1.0])

    # =========================================================================
    # PAGE 9: BLANK SPACER PAGE
    # =========================================================================
    doc.add_page_break()
    p_blank = doc.add_paragraph()
    p_blank.paragraph_format.space_before = Pt(280)
    p_blank.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_b = p_blank.add_run("")

    # =========================================================================
    # PAGE 10: DIVIDER - SYNOPSIS
    # =========================================================================
    doc.add_page_break()
    add_divider_page("SYNOPSIS")

    # =========================================================================
    # PAGE 11: SYNOPSIS CONTENT
    # =========================================================================
    doc.add_page_break()
    p_syn = doc.add_paragraph()
    p_syn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_syn.paragraph_format.space_before = Pt(20)
    p_syn.paragraph_format.space_after = Pt(16)
    r_syn = p_syn.add_run("SYNOPSIS")
    r_syn.font.name = 'Times New Roman'
    r_syn.font.size = Pt(14)
    r_syn.bold = True

    add_p("The Online Examination System using Python is a web-based software application developed to manage and automate academic examinations, student assessments, and institutional evaluations through a secure centralized system. In traditional educational institutions, conducting examinations requires extensive manual coordination, paper printing, physical hall invigilation, manual script grading, and manual ledger record-keeping, which are time-consuming and prone to human calculation errors.")
    add_p("This system provides a centralized platform for managing student details, exam configurations, question banks, candidate attempts, auto-grading, and integrity auditing. The application allows administrators to author and publish multi-format exams, manage questions, evaluate student submissions, and analyze class performance. Students can securely log in to view available tests, take examinations within timed intervals, view real-time question statuses, submit responses, and review comprehensive scorecards and verifiable digital certificates.")
    add_p("A key security highlight of the system is its server-authoritative timer and proctoring audit log. The remaining duration is calculated independently on the backend server, completely preventing client-side timer tampering. Simultaneously, client browser events such as window blur, tab switching, and clipboard infractions are logged in real time. An asynchronous background auto-save mechanism permanently persists candidate choices without page reloads.")
    add_p("The system is developed using Python Flask as the backend web framework, SQLite as the transactional relational database, and HTML5, CSS3, and JavaScript for the responsive user interface. The primary objective of this project is to eliminate paper consumption, streamline grading, maintain exam security, provide actionable feedback to learners, and deliver a modern, efficient assessment experience for administrators, faculty members, and students.")

    # =========================================================================
    # PAGE 12: DIVIDER - INTRODUCTION
    # =========================================================================
    doc.add_page_break()
    add_divider_page("INTRODUCTION")

    # =========================================================================
    # PAGE 13: 1. INTRODUCTION
    # =========================================================================
    doc.add_page_break()
    add_page_title("1. INTRODUCTION")
    add_p("The Online Examination System using Python is a web-based software platform designed to manage institutional academic assessments, candidate evaluations, and automated grading in an efficient, secure, and organized manner. In many colleges and universities, examination processes are still conducted manually or through fragmented survey tools, which require significant physical labor, high paper expenditure, and extensive manual verification.")
    add_p("This system provides a centralized web platform for managing students, examinations, multi-format question banks, candidate attempts, and integrity logs. The administrator can create examinations, specify time limits, set passing thresholds, compose multiple-choice or short-answer questions, and analyze class scores. Students can securely log in to access scheduled tests, track their remaining time, select answers with instant background persistence, and view performance scorecards upon completion.")
    add_p("The platform also provides automated proctoring hooks that monitor candidate tab switches and window focus losses, discouraging academic dishonesty during live testing sessions. Furthermore, objective questions are graded instantly by the backend engine, eliminating human grading delays and providing examinees with immediate clarity regarding their concept mastery.")
    add_p("The main purpose of this project is to simplify institutional examination management, eliminate paper wastage, improve assessment security through server-authoritative timing, reduce manual administrative workload, and provide reliable, real-time academic records for students, educators, and college administrators.")

    # =========================================================================
    # PAGE 14: 1.1 SYSTEM SPECIFICATION & HARDWARE
    # =========================================================================
    doc.add_page_break()
    add_page_title("1.1 SYSTEM SPECIFICATION")
    add_p("The system specification describes the hardware and software requirements needed to develop, execute, host, and access the Online Examination System using Python. The system utilizes standard computer hardware, modern web browsers, the Python runtime, the Flask web framework, SQLite database, and supporting client-side technologies.")

    add_subheading("1.1.1 HARDWARE CONFIGURATION")
    add_p("The hardware requirements for server hosting and client workstations are:")
    add_bullet("Intel Core i3 processor or higher (2.0 GHz or above). A multi-core Intel Core i5/i7 processor is recommended for servers hosting high-concurrency examination halls.", "Processor: ")
    add_bullet("4 GB or above (8 GB recommended for concurrent session handling and database operations).", "RAM: ")
    add_bullet("500 GB or above HDD / 256 GB SSD with at least 5 GB free disk space for application files, logs, and database records.", "Hard Disk: ")
    add_bullet("Standard Color LED/LCD Monitor supporting 1366x768 or 1920x1080 resolution.", "Monitor: ")
    add_bullet("Standard 104-key USB/Wireless Keyboard.", "Keyboard: ")
    add_bullet("Standard USB/Wireless Optical Mouse.", "Mouse: ")
    add_bullet("Broadband LAN or Wi-Fi connection (minimum 2 Mbps per client) for real-time exam communication and background response auto-saving.", "Internet / Network: ")
    add_bullet("Standard Desktop PC, Laptop, or Tablet capable of running a modern web browser.", "Client Node: ")

    add_subheading("1.1.2 SOFTWARE CONFIGURATION")
    add_p("The core software environment consists of:")
    add_bullet("Microsoft Windows 10 / Windows 11 (64-bit), or Linux (Ubuntu 20.04 LTS+).", "Operating System: ")
    add_bullet("Python 3.10 to 3.13.", "Programming Language: ")
    add_bullet("Flask (Version 3.0.x).", "Web Framework: ")
    add_bullet("SQLite 3 (Transactional Relational Database).", "Database: ")
    add_bullet("HTML5, CSS3, JavaScript (ES6+).", "Front-End: ")
    add_bullet("Google Chrome / Microsoft Edge / Mozilla Firefox.", "Web Browser: ")
    add_bullet("Visual Studio Code (VS Code).", "Code Editor / IDE: ")
    add_bullet("Flask Development Server / Waitress WSGI Server.", "Web Server: ")

    # =========================================================================
    # PAGE 15: 1.1.2 SOFTWARE CONFIGURATION (Detailed Description Part 1)
    # =========================================================================
    doc.add_page_break()
    add_page_title("1.1.2 SOFTWARE CONFIGURATION (Detailed Software Technologies)")
    add_p("The Online Examination System is developed using a cohesive set of software technologies to provide a high-performance, secure, and user-friendly testing platform. Each software component performs a specialized role in backend processing, data persistence, and interface delivery.")

    add_subheading("1. Python")
    add_p("Python is a high-level, interpreted programming language known for its readability, robust security features, and extensive standard library. In this application, Python forms the backbone of the backend server. It executes business logic, enforces role-based access control, computes server-authoritative time limits, calculates proportional partial credit for multiple-choice questions, and coordinates database transactions.")

    add_subheading("2. Flask")
    add_p("Flask is a lightweight, modular web framework for Python. It provides core routing, HTTP request handling, session state management, and template rendering through Jinja2. Flask's Blueprint architecture allows the application to be cleanly partitioned into discrete functional modules—such as authentication, student testing, admin management, and REST APIs—ensuring maintainability and testability.")

    add_subheading("3. SQLite")
    add_p("SQLite is a self-contained, serverless, zero-configuration relational database management system. It stores all critical institutional records within a compact transactional file (`oes.db`). SQLite maintains tables for user credentials, examination configurations, question banks, student attempts, individual answers, and proctoring audit events with complete ACID compliance.")

    add_subheading("4. HTML5")
    add_p("HTML5 (HyperText Markup Language) defines the semantic structure of all web pages across the platform. It provides accessible form fields, navigation hierarchies, interactive question palettes, and responsive layouts for dashboards, exam interfaces, scorecards, and digital certificates.")

    # =========================================================================
    # PAGE 16: 1.1.2 SOFTWARE CONFIGURATION (Detailed Description Part 2)
    # =========================================================================
    doc.add_page_break()
    add_page_title("1.1.2 SOFTWARE CONFIGURATION (Frontend & Tooling)")

    add_subheading("5. CSS3 (Cascading Style Sheets)")
    add_p("CSS3 is used to design, layout, and style the web application. The platform incorporates a custom dark-mode design system utilizing CSS variables, responsive CSS grids, flexbox layouts, and glassmorphism cards. CSS controls visual feedback for question states (answered, unanswered, flagged), status badges, and typography across devices.")

    add_subheading("6. JavaScript (ES6+)")
    add_p("Vanilla JavaScript makes the examination interface dynamic, responsive, and resilient. It runs real-time countdown clocks, handles question navigation, sends asynchronous Fetch API requests to auto-save candidate answers without page reloads, and binds browser event listeners for proctoring detection.")

    add_subheading("7. Web Browser")
    add_p("A modern standards-compliant web browser (such as Google Chrome or Microsoft Edge) serves as the primary client execution runtime. It renders HTML/CSS, executes JavaScript client scripts, and interfaces with the backend server via HTTP/HTTPS protocols.")

    add_subheading("8. Visual Studio Code")
    add_p("Visual Studio Code (VS Code) is the primary integrated development environment (IDE) utilized for authoring, debugging, and maintaining the Python, HTML, CSS, and JavaScript source files. It provides syntax highlighting, linting, Git integration, and debugging tools.")

    add_subheading("9. Security & Anti-Tamper Engine")
    add_p("The backend implements cryptographic password hashing using Werkzeug's PBKDF2-SHA256 algorithm. Session cookies are protected with HttpOnly and SameSite flags, and sensitive routes are protected by server-side role verification decorators (`@login_required`, `@admin_required`).")

    # =========================================================================
    # PAGE 17: 1.1.2 SOFTWARE CONFIGURATION (Proctoring & Summary)
    # =========================================================================
    doc.add_page_break()
    add_page_title("1.1.2 SOFTWARE CONFIGURATION (Proctoring & Summary)")

    add_subheading("10. Proctoring & Event Monitor")
    add_p("The system includes client-side hooks that monitor candidate window visibility. If a student minimizes the browser, switches tabs, or attempts to copy question text to the clipboard, JavaScript intercepts the event and sends an asynchronous alert to `/api/integrity-event`. The backend logs the infraction timestamp, event type, and candidate details, adjusting the session's overall integrity score.")

    add_subheading("Software Technology Integration Summary")
    add_p("Together, these software technologies create an integrated, robust, and scalable testing environment. The combination of Python Flask, SQLite, HTML5, CSS3, and JavaScript ensures:")
    add_bullet("Zero client-side timer manipulation through backend time enforcement.", "Reliable Testing: ")
    add_bullet("Permanent answer persistence through asynchronous auto-saving.", "Resilient Sessions: ")
    add_bullet("Instant grade dissemination upon test submission.", "Efficient Grading: ")
    add_bullet("Actionable candidate oversight via real-time audit logging.", "Institutional Security: ")
    add_bullet("A modern, accessible user experience across desktop and mobile devices.", "High Usability: ")

    # =========================================================================
    # PAGE 18: DIVIDER - SYSTEM STUDY
    # =========================================================================
    doc.add_page_break()
    add_divider_page("SYSTEM STUDY")

    # =========================================================================
    # PAGE 19: 2. SYSTEM STUDY
    # =========================================================================
    doc.add_page_break()
    add_page_title("2. SYSTEM STUDY")
    add_p("System Study is a foundational phase in software development. It involves analyzing the existing examination methodologies, identifying structural limitations and operational costs, collecting user requirements from faculty and students, and engineering a superior proposed system. The main objective of the system study is to understand the pedagogical and operational needs of educational institutions and determine how software automation can eliminate inefficiencies.")
    add_p("The Online Examination System using Python is developed to integrate examination scheduling, question authoring, live testing, automated evaluation, and integrity monitoring into a single unified web application. The platform provides distinct, secure workspaces for Administrators, Examinees, and Executive Department Heads.")
    add_p("By digitizing the assessment lifecycle, the system eliminates paper consumption, reduces administrative turnaround time, prevents client-side timer fraud, and provides instant performance analytics. Faculty members can author rich question banks with difficulty tiers, while students gain immediate visibility into their grades, explanations, and downloadable certificates.")

    # =========================================================================
    # PAGE 20: 2.1 EXISTING SYSTEM
    # =========================================================================
    doc.add_page_break()
    add_page_title("2.1 EXISTING SYSTEM")
    add_subheading("2.1.1 DESCRIPTION")
    add_p("In the existing conventional system, academic examinations are conducted through manual paper-based processes or basic survey forms (such as Google Forms). Faculty members manually draft question papers, submit them for physical printing, and store question packets under physical security until exam day. Students sit in physical halls where invigilators distribute papers and track time using room clocks. Completed scripts are collected, sorted, and delivered to instructors for manual grading. Scores are then transcribed into spreadsheets or physical registers.")

    add_subheading("2.1.2 DRAWBACKS")
    add_p("The major drawbacks of the existing manual examination system are:")
    add_bullet("Manual paper distribution and script collection require significant staff time and effort.")
    add_bullet("Manual grading and marks entry are highly vulnerable to human calculation errors.")
    add_bullet("Physical question paper printing and answer booklet storage incur substantial financial and environmental costs.")
    add_bullet("Compiling marks registers and calculating class statistics takes weeks to complete.")
    add_bullet("Basic online forms rely on client-side timers that can be bypassed by stopping browser JavaScript.")
    add_bullet("Standard web forms lack background auto-save, resulting in total answer loss if a browser reloads.")
    add_bullet("Paper and standard forms lack automated proctoring to detect tab switching or web searching.")
    add_bullet("Scores are provided without detailed concept-level explanations or personalized study feedback.")
    add_bullet("Searching past examination records requires manual inspection of physical file cabinets.")
    add_bullet("Generating verifiable academic certificates requires manual drafting and physical signing.")
    add_bullet("Communication between examiners, administrators, and students is fragmented and delayed.")
    add_bullet("There is no centralized digital platform uniting question banking, live testing, and audit logging.")

    # =========================================================================
    # PAGE 21: 2.2 PROPOSED SYSTEM
    # =========================================================================
    doc.add_page_break()
    add_page_title("2.2 PROPOSED SYSTEM")
    add_p("The proposed Online Examination System using Python is a web-based platform developed to overcome the limitations of the existing manual and fragmented assessment systems.")
    add_p("The proposed system provides a centralized transactional database for storing and managing user profiles, examinations, questions, student attempt sessions, individual answers, and proctoring audit events. It provides dedicated, role-protected workspaces for Administrators, Students, and Executive Department Heads.")

    add_subheading("2.2.1 DESCRIPTION")
    add_p("The proposed system consists of several integrated operational modules:")

    add_subheading("Admin Module:")
    add_p("The administrator can securely log in to access the administrative dashboard. The admin can create new examinations, configure duration and pass thresholds, author questions across four distinct formats, manage question banks, publish exams, and inspect aggregate class analytics.")

    add_subheading("Student Module:")
    add_p("Students log in using their verified email and password. They can view scheduled and published assessments, initiate timed test attempts, navigate through question palettes, select answers with automatic background persistence, and access finalized scorecards and digital certificates.")

    # =========================================================================
    # PAGE 22: 2.2.1 DESCRIPTION (Continued)
    # =========================================================================
    doc.add_page_break()
    add_page_title("2.2.1 DESCRIPTION (Continued)")

    add_subheading("Exam Authoring & Question Bank Module:")
    add_p("Enables instructors to create structured question items with custom point values, difficulty ratings (Easy, Medium, Hard), topic tags, and detailed answer explanations. Supports single-choice, multiple-choice, true/false, and short-answer formats.")

    add_subheading("Server-Authoritative Timer Module:")
    add_p("Calculates remaining time independently on the backend server. The server establishes an immutable target completion time (`end_time`). Even if an examinee alters their device clock, pauses JavaScript, or reloads the browser, the server strictly enforces the deadline and automatically finalizes the exam when time expires.")

    add_subheading("Asynchronous Auto-Save Module:")
    add_p("Uses asynchronous JavaScript Fetch API calls to transmit candidate responses to `/api/save-answer` in real time. Answers are written directly to SQLite without requiring page refreshes, providing complete protection against sudden power outages or network drops.")

    add_subheading("Proctoring & Audit Module:")
    add_p("Monitors client-side visibility events. If an examinee navigates away from the examination window, switches tabs, or copies question content, an infraction event is logged with a precise timestamp. Department Heads can inspect the live proctoring log to identify suspicious activity.")

    add_subheading("Automated Grading Module:")
    add_p("Evaluates objective question submissions immediately upon completion. Implements exact matching for single-choice and true/false questions, proportional partial credit for multiple-choice combinations, and normalized string matching for short-answer questions.")

    # =========================================================================
    # PAGE 23: 2.2.2 FEATURES
    # =========================================================================
    doc.add_page_break()
    add_page_title("2.2.2 FEATURES")
    add_p("The major features of the proposed Online Examination System are:")
    add_bullet("1. Admin Login and Dashboard – Centralized workspace for managing institutional assessment operations.")
    add_bullet("2. Role-Based Access Control – Strict permission separation between Students, Admins, and Department Heads.")
    add_bullet("3. Exam Management – Tools to create, edit, configure, and publish assessments with custom rules.")
    add_bullet("4. Multi-Format Question Authoring – Support for single-choice, multi-choice, true/false, and short-answer items.")
    add_bullet("5. Question Bank Classification – Categorizes questions by academic topic and difficulty levels.")
    add_bullet("6. Server-Authoritative Timer – Backend calculation of remaining time preventing client-side tampering.")
    add_bullet("7. Real-Time Auto-Save – Asynchronous background persistence of student responses without page reloads.")
    add_bullet("8. Question Status Palette – Interactive visual grid indicating answered, unanswered, and current questions.")
    add_bullet("9. Automated Instant Grading – Immediate calculation of scores and percentages upon submission.")
    add_bullet("10. Proportional Partial Credit – Fair scoring algorithm for multi-choice questions with multiple correct options.")
    add_bullet("11. Client-Side Proctoring Hooks – Automated detection of tab switching, window blur, and copy attempts.")
    add_bullet("12. Executive Proctoring Audit Log – Real-time chronological audit trail of candidate infractions.")
    add_bullet("13. Comprehensive Scorecards – Question-by-question review displaying correct answers and educational rationales.")
    add_bullet("14. Verifiable Digital Certificates – Automatically generates achievement certificates with unique verification IDs.")
    add_bullet("15. Performance Analytics – Class-wide metric cards, score distribution graphs, and pass rate tracking.")
    add_bullet("16. Student Dashboard – Personalized portal listing available assessments, attempt histories, and certificates.")
    add_bullet("17. Executive User Management – Department Head tools to view, promote, or update user account roles.")
    add_bullet("18. Password Security – Salted PBKDF2-SHA256 password hashing protecting student and faculty accounts.")
    add_bullet("19. Centralized Transactional Database – ACID-compliant SQLite storage maintaining all system records.")
    add_bullet("20. Responsive Modern Interface – Glassmorphism-styled web pages offering intuitive desktop and mobile access.")

    # =========================================================================
    # PAGE 24: ADVANTAGES OF THE PROPOSED SYSTEM
    # =========================================================================
    doc.add_page_break()
    add_page_title("ADVANTAGES OF THE PROPOSED SYSTEM")
    add_p("The proposed Online Examination System provides significant operational, security, and pedagogical advantages over the existing manual system:")
    add_bullet("Eliminates paper, printing, transportation, and physical storage costs.", "1) Paperless Efficiency: ")
    add_bullet("Automated grading delivers instant results, saving hundreds of hours of manual script evaluation.", "2) Instant Evaluation: ")
    add_bullet("Backend timer calculation prevents candidates from extending exam time through local clock manipulation.", "3) Tamper-Proof Timing: ")
    add_bullet("Asynchronous background auto-saving guarantees that candidate answers are never lost due to network or browser crashes.", "4) Zero Data Loss: ")
    add_bullet("Automated tab-switch and window-blur tracking discourages academic dishonesty during examinations.", "5) Active Proctoring: ")
    add_bullet("Examinees receive instant scorecards complete with correct answer explanations and topic feedback.", "6) Actionable Feedback: ")
    add_bullet("Passing examinees receive an official digital certificate featuring a unique cryptographic ID for instant verification.", "7) Verifiable Credentials: ")
    add_bullet("Provides dedicated, role-tailored dashboards for Students, Faculty Admins, and Executive Department Heads.", "8) Role-Tailored Access: ")
    add_bullet("Centralizes questions, student submissions, scores, and audit records within a secure SQLite database.", "9) Centralized Archiving: ")
    add_bullet("Provides an elegant, responsive interface that functions smoothly across computers, laptops, and tablets.", "10) Modern User Experience: ")

    # =========================================================================
    # PAGE 25: DIVIDER - SYSTEM DESIGN AND DEVELOPMENT
    # =========================================================================
    doc.add_page_break()
    add_divider_page("SYSTEM DESIGN AND DEVELOPMENT")

    # =========================================================================
    # PAGE 26: 3. SYSTEM DESIGN AND DEVELOPMENT & 3.1 FILE DESIGN
    # =========================================================================
    doc.add_page_break()
    add_page_title("3. SYSTEM DESIGN AND DEVELOPMENT")
    add_p("System Design and Development describes the architectural structure, file organization, inputs, outputs, database schema, and coding methods used in the implementation of the Online Examination System using Python.")
    add_p("The system is engineered as a full-stack web application with a Python Flask backend and an SQLite transactional database. The design emphasizes modularity, separation of concerns, robust security, and seamless user interaction.")

    add_subheading("3.1 FILE DESIGN")
    add_p("File design describes the directory structure, configuration files, scripts, and modular components of the application. The system uses an SQLite database (`data/oes.db`) to manage all persistent records.")
    add_p("The project root contains key operational scripts:")
    add_bullet("`run.py` – Application entry script that instantiates and launches the Flask web server.")
    add_bullet("`config.py` – Central configuration file defining secret keys, database URIs, and session parameters.")
    add_bullet("`seed.py` – Database seeding utility that initializes default admin, head, and student accounts along with sample assessments.")

    # =========================================================================
    # PAGE 27: 3.1 FILE DESIGN (Continued - Blueprints & Models)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.1 FILE DESIGN (Blueprints & Core Packages)")
    add_p("The core codebase resides within the `app/` package, organized into specialized subdirectories following clean architectural principles:")

    add_subheading("Application Factory (`app/__init__.py`)")
    add_p("Initializes the Flask application, configures SQLAlchemy ORM, binds Flask-Login session authentication, and registers modular blueprints for URL dispatching.")

    add_subheading("Data Models (`app/models/schema.py`)")
    add_p("Defines object-relational mapping classes: `User`, `Exam`, `Question`, `Attempt`, `Answer`, `IntegrityEvent`, and `AIFeedback`. Each model encapsulates database columns, constraints, foreign keys, and helper serialization methods (`to_dict`).")

    add_subheading("Route Controllers (`app/routes/`)")
    add_bullet("`auth.py` – Manages registration, login validation, session tracking, and logout.")
    add_bullet("`student.py` – Handles student dashboard views, exam taking, timer countdowns, results, and certificate rendering.")
    add_bullet("`admin.py` – Handles exam creation, publishing, question bank authoring, and analytics.")
    add_bullet("`head.py` – Handles executive user management and proctoring audit log inspection.")
    add_bullet("`api.py` – Provides REST endpoints for asynchronous response saving and integrity event reporting.")

    # =========================================================================
    # PAGE 28: 3.1 FILE DESIGN (Continued - Services & Static Assets)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.1 FILE DESIGN (Services & Static Assets)")

    add_subheading("Business Logic Services (`app/services/`)")
    add_bullet("`scoring.py` – Evaluates student attempts, computes exact-match and proportional partial credit scores, and determines pass/fail status.")
    add_bullet("`vector_service.py` – Integrates ChromaDB vector search for question semantic retrieval and deduplication.")
    add_bullet("`ai_service.py` – Coordinates AI API interactions for automated question generation and candidate performance summaries.")

    add_subheading("Static Assets (`app/static/`)")
    add_bullet("`css/style.css` – Modern dark-mode glassmorphic styling system, CSS custom properties, responsive grids, and question status indicators.")
    add_bullet("`js/exam.js` – Client-side script handling server timer synchronization, auto-save Fetch API calls, and visibility change listeners.")

    add_subheading("HTML Templates (`app/templates/`)")
    add_bullet("`base.html` – Master template containing header, navigation bar, flash messages, and footer.")
    add_bullet("Subdirectories `auth/`, `admin/`, `student/`, and `head/` containing modular Jinja2 templates for each user role.")

    # =========================================================================
    # PAGE 29: 3.1 FILE DESIGN (Summary Table)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.1 FILE DESIGN (File Structure Summary)")
    add_p("The complete file architecture of the Online Examination System is summarized below:")

    f_headers = ["File / Path", "Role / Type", "Functional Responsibility"]
    f_rows = [
        ["run.py", "Server Entry", "Launches Flask development/production WSGI server."],
        ["config.py", "Configuration", "Stores database URIs, secret keys, session lifetime."],
        ["seed.py", "Database Seeder", "Seeds default accounts and demo assessment items."],
        ["app/__init__.py", "App Factory", "Initializes extensions and registers all route blueprints."],
        ["app/models/schema.py", "Data Schema", "Declares SQLAlchemy models, tables, and relationships."],
        ["app/routes/auth.py", "Controller", "Manages user authentication and role access."],
        ["app/routes/student.py", "Controller", "Renders exam taking, timer sync, and scorecards."],
        ["app/routes/admin.py", "Controller", "Manages exam creation, question bank, and analytics."],
        ["app/routes/head.py", "Controller", "Executive dashboard, user administration, proctoring log."],
        ["app/routes/api.py", "REST API", "Asynchronous auto-save and integrity event endpoints."],
        ["app/services/scoring.py", "Service", "Executes automated grading and partial credit logic."],
        ["app/static/css/style.css", "Design System", "CSS custom properties, glassmorphic UI, responsive grids."],
        ["app/static/js/exam.js", "Client Engine", "Timer countdown, visibility listeners, async save calls."],
        ["app/templates/base.html", "Base View", "Unified layout, navigation bars, dynamic user badges."]
    ]
    add_table_data(f_headers, f_rows, col_widths=[1.8, 1.4, 3.3])

    # =========================================================================
    # PAGE 30: 3.2 INPUT DESIGN
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.2 INPUT DESIGN")
    add_p("Input design describes how data enters the system from users. Proper input design minimizes errors, enforces data validation, and ensures an intuitive user experience.")

    add_subheading("User Authentication Inputs")
    add_bullet("Name: Candidate full legal name (alphabetical string, 2 to 120 chars).")
    add_bullet("Email: Validated RFC-compliant institutional email address (unique login ID).")
    add_bullet("Password: Minimum 6 characters, masked input, hashed via PBKDF2 before storage.")
    add_bullet("Role: User permission tier (`student`, `admin`, or `head`).")

    add_subheading("Exam Configuration Inputs")
    add_bullet("Exam Title: Descriptive title of the assessment (e.g., 'Python Fundamentals').")
    add_bullet("Description: Guidelines, syllabus topics, and candidate instructions.")
    add_bullet("Duration: Total allowable examination duration in minutes (5 to 180 min).")
    add_bullet("Pass Mark: Minimum passing score percentage (0.0 to 100.0%).")
    add_bullet("Published State: Boolean switch toggling student accessibility.")

    # =========================================================================
    # PAGE 31: 3.2 INPUT DESIGN (Continued)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.2 INPUT DESIGN (Continued)")

    add_subheading("Question Creation Inputs")
    add_bullet("Question Text: Full query text stating the problem clearly.")
    add_bullet("Question Type: Selector for `single_choice`, `multi_choice`, `true_false`, or `short_answer`.")
    add_bullet("Options List: Dynamic array of selectable choices for multiple-choice items.")
    add_bullet("Correct Answer(s): Key value or array of correct option identifiers.")
    add_bullet("Points: Marks weightage allocated to the question (Float >= 0.5).")
    add_bullet("Topic: Academic subject matter classification (e.g., 'Loops', 'OOP').")
    add_bullet("Difficulty: Complexity rating (`easy`, `medium`, `hard`).")
    add_bullet("Explanation: Educational justification displayed after grading.")

    add_subheading("Examination Response Inputs (Live Testing)")
    add_bullet("Attempt ID: Unique attempt session identifier.")
    add_bullet("Question ID: Specific question item being answered.")
    add_bullet("Student Response: Selected radio value, checkbox array, or text response.")

    add_subheading("Proctoring Anomaly Inputs")
    add_bullet("Event Type: Anomaly code (`tab_switch`, `window_blur`, `copy_attempt`).")
    add_bullet("Details: Contextual description of the detected browser event.")

    # =========================================================================
    # PAGE 32: 3.3 OUTPUT DESIGN
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.3 OUTPUT DESIGN")
    add_p("Output design describes the information displayed by the system after processing user actions. Outputs are designed to be clear, actionable, and visually appealing.")

    add_subheading("Admin Dashboard Outputs")
    add_bullet("Total Exams: Aggregate count of created and published assessments.")
    add_bullet("Question Bank Size: Total number of active assessment questions.")
    add_bullet("Total Attempts: Number of student test attempts recorded.")
    add_bullet("Average Class Score: Global performance percentage across all assessments.")

    add_subheading("Student Dashboard Outputs")
    add_bullet("Available Assessments: Grid cards showing active exams, duration, and pass thresholds.")
    add_bullet("Assessment History: Table listing past attempts, percentage scores, and passed/failed badges.")
    add_bullet("Action Buttons: Direct links to begin exam, review completed attempts, or download certificates.")

    # =========================================================================
    # PAGE 33: 3.3 OUTPUT DESIGN (Continued - Live Exam Interface)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.3 OUTPUT DESIGN (Live Exam Interface)")

    add_subheading("Live Examination Screen Outputs")
    add_bullet("Real-Time Countdown Timer: Prominent visual clock displaying remaining minutes and seconds, synchronized with the server-authoritative end timestamp.")
    add_bullet("Auto-Save Status Indicator: Real-time badge confirming whether the latest answer selection has been persisted to the server ('Saving...', 'All responses saved').")
    add_bullet("Question Navigation Palette: Grid showing question numbers color-coded by status (Answered = Green, Unanswered = Gray, Active = Blue outline).")
    add_bullet("Exam Details Banner: Title, candidate name, total questions, and maximum possible marks.")
    add_bullet("Submit Assessment Button: Final submission trigger with confirmation modal prompt.")

    # =========================================================================
    # PAGE 34: 3.3 OUTPUT DESIGN (Continued - Reports & Certificates)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.3 OUTPUT DESIGN (Reports & Certificates)")

    add_subheading("Scorecard & Performance Outputs")
    add_bullet("Final Score: Marks earned out of total points (e.g., 26.0 / 30.0 pts).")
    add_bullet("Score Percentage: Calculated percentage vs. pass threshold (e.g., 86.7% vs. 60.0%).")
    add_bullet("Pass/Fail Indicator: High-visibility badge ('PASSED' in green or 'NEEDS IMPROVEMENT' in red).")
    add_bullet("Session Integrity Rating: Calculated integrity percentage reflecting infractions logged during testing.")
    add_bullet("Question-by-Question Review: Displays candidate's answer, correct answer key, points awarded, and explanation.")

    add_subheading("Digital Certificate of Achievement")
    add_bullet("Official completion certificate featuring institution name, candidate name, examination title, score, date, and unique verification ID.")

    add_subheading("Executive Proctoring Audit Log Output")
    add_bullet("Chronological table displaying timestamp, candidate email, attempt ID, infraction type, and resulting integrity rating.")

    # =========================================================================
    # PAGE 35: 3.4 DATABASE DESIGN (Schema Tables 1 & 2)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.4 DATABASE DESIGN")
    add_p("The database is implemented in SQLite using third normal form (3NF) principles to ensure data integrity and eliminate data redundancy.")

    add_subheading("Table 1: users (User Account Credentials & Roles)")
    u_h = ["Field Name", "Data Type", "Constraint", "Description"]
    u_r = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique user ID."],
        ["name", "VARCHAR(120)", "NOT NULL", "Full legal name."],
        ["email", "VARCHAR(120)", "NOT NULL, UNIQUE, INDEX", "Institutional email."],
        ["password_hash", "VARCHAR(256)", "NOT NULL", "PBKDF2 salted hash."],
        ["role", "VARCHAR(20)", "NOT NULL, DEFAULT 'student'", "student, admin, or head."],
        ["created_at", "DATETIME", "DEFAULT UTC NOW", "Registration date."]
    ]
    add_table_data(u_h, u_r, col_widths=[1.5, 1.4, 2.0, 1.6])

    add_subheading("Table 2: exams (Examination Definitions)")
    e_h = ["Field Name", "Data Type", "Constraint", "Description"]
    e_r = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique exam ID."],
        ["title", "VARCHAR(200)", "NOT NULL", "Exam title."],
        ["description", "TEXT", "NULLABLE", "Exam instructions."],
        ["duration_minutes", "INTEGER", "NOT NULL, DEFAULT 30", "Allowed time (minutes)."],
        ["pass_mark", "FLOAT", "NOT NULL, DEFAULT 50.0", "Pass percentage."],
        ["is_published", "BOOLEAN", "DEFAULT FALSE", "Visibility flag."],
        ["created_by_id", "INTEGER", "FOREIGN KEY -> users.id", "Admin creator ID."],
        ["created_at", "DATETIME", "DEFAULT UTC NOW", "Creation timestamp."]
    ]
    add_table_data(e_h, e_r, col_widths=[1.5, 1.4, 2.0, 1.6])

    # =========================================================================
    # PAGE 36: 3.4 DATABASE DESIGN (Schema Tables 3, 4, 5, 6)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.4 DATABASE DESIGN (Continued)")

    add_subheading("Table 3: questions (Assessment Items)")
    q_h = ["Field Name", "Data Type", "Constraint", "Description"]
    q_r = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique question ID."],
        ["exam_id", "INTEGER", "FOREIGN KEY -> exams.id", "Exam foreign key."],
        ["text", "TEXT", "NOT NULL", "Question content text."],
        ["question_type", "VARCHAR(30)", "NOT NULL", "single_choice, multi_choice, etc."],
        ["options_json", "TEXT", "NULLABLE", "JSON options array."],
        ["correct_answers_json", "TEXT", "NOT NULL", "JSON answers array."],
        ["points", "FLOAT", "NOT NULL, DEFAULT 1.0", "Marks allocated."]
    ]
    add_table_data(q_h, q_r, col_widths=[1.5, 1.4, 2.0, 1.6])

    add_subheading("Table 4: attempts (Student Exam Sessions)")
    a_h = ["Field Name", "Data Type", "Constraint", "Description"]
    a_r = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Attempt session ID."],
        ["user_id", "INTEGER", "FOREIGN KEY -> users.id", "Student ID."],
        ["exam_id", "INTEGER", "FOREIGN KEY -> exams.id", "Exam ID."],
        ["start_time", "DATETIME", "NOT NULL", "Server start timestamp."],
        ["end_time", "DATETIME", "NOT NULL", "Server-authoritative end time."],
        ["status", "VARCHAR(30)", "NOT NULL", "in_progress, submitted, graded."],
        ["score", "FLOAT", "DEFAULT 0.0", "Final score achieved."],
        ["integrity_score", "INTEGER", "DEFAULT 100", "Proctoring score (100-0)."]
    ]
    add_table_data(a_h, a_r, col_widths=[1.5, 1.4, 2.0, 1.6])

    # =========================================================================
    # PAGE 37: 3.5 SYSTEM DEVELOPMENT
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.5 SYSTEM DEVELOPMENT")
    add_p("System development is the engineering process of translating the structural and database designs into a functioning web application. The Online Examination System was built using Python Flask following an iterative Model-View-Controller (MVC) architectural design:")

    add_subheading("Model Layer (Data & Business Entities)")
    add_p("Implemented using Flask-SQLAlchemy in `app/models/schema.py`. The models declare database tables, column types, relational constraints, foreign keys, and instance methods (such as `Attempt.remaining_seconds()` which computes backend time remaining).")

    add_subheading("View Layer (User Interface & Presentation)")
    add_p("Implemented using Jinja2 templates in `app/templates/` styled with modern CSS3 variables in `app/static/css/style.css`. The view layer renders responsive dashboards, clean question cards, and accessible scorecards.")

    add_subheading("Controller Layer (Request Handling & Routing)")
    add_p("Organized across discrete Flask Blueprints in `app/routes/`. Each blueprint processes incoming HTTP requests, validates form submissions, queries data models, invokes services, and returns rendered HTML or JSON responses.")

    # =========================================================================
    # PAGE 38: 3.6 DESCRIPTION OF MODULES (Auth & Student)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.6 DESCRIPTION OF MODULES")

    add_subheading("3.6.1 AUTHENTICATION MODULE (app/routes/auth.py)")
    add_p("The Authentication Module manages user onboarding, credential verification, and session state. It provides:")
    add_bullet("User Registration: Validates name, institutional email uniqueness, and password criteria before creating student accounts.")
    add_bullet("Secure Login: Authenticates user credentials using PBKDF2 password hash comparisons, initializing Flask-Login user sessions.")
    add_bullet("Role Redirection: Directs authenticated users to role-specific dashboards (Admin, Head, or Student).")
    add_bullet("Session Teardown: Clears session state and invalidates cookies upon logout.")

    add_subheading("3.6.2 STUDENT EXAMINATION MODULE (app/routes/student.py)")
    add_p("The Student Module manages the examinee experience throughout the assessment lifecycle:")
    add_bullet("Dashboard: Lists active published examinations and past attempt histories.")
    add_bullet("Attempt Initialization: Creates a new `Attempt` row and calculates the server-authoritative `end_time`.")
    add_bullet("Exam Execution: Serves the live exam view with the real-time countdown timer and question palette.")
    add_bullet("Scorecard & Review: Displays finalized marks, pass/fail status, and question-by-question rationales.")
    add_bullet("Certificate Generator: Renders the official academic certificate of achievement for passing attempts.")

    # =========================================================================
    # PAGE 39: 3.6 DESCRIPTION OF MODULES (Admin & Head)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.6 DESCRIPTION OF MODULES (Admin & Executive)")

    add_subheading("3.6.3 ADMINISTRATOR MODULE (app/routes/admin.py)")
    add_p("The Administrator Module provides educators with full assessment management capabilities:")
    add_bullet("Exam Creation & Editing: Configures titles, syllabus descriptions, time limits, and passing thresholds.")
    add_bullet("Question Bank Management: Authors single-choice, multi-choice, true/false, and short-answer items with point values, topic tags, and difficulty tiers.")
    add_bullet("Exam Publishing: Toggles draft assessments into published state, making them accessible to students.")
    add_bullet("Class Analytics: Aggregates student completion counts, average scores, and question pass rates.")

    add_subheading("3.6.4 EXECUTIVE HEAD MODULE (app/routes/head.py)")
    add_p("The Executive Head Module provides institutional oversight for Department Heads:")
    add_bullet("User Management: View all registered users across the institution and promote/modify user roles.")
    add_bullet("Proctoring Audit Log: Inspects live stream of integrity infractions (tab switches, window defocus) across all candidates.")
    add_bullet("Executive Dashboard: Monitors high-level metrics including total candidates, exams, and compliance ratings.")

    # =========================================================================
    # PAGE 40: 3.6 DESCRIPTION OF MODULES (API & Services)
    # =========================================================================
    doc.add_page_break()
    add_page_title("3.6 DESCRIPTION OF MODULES (API & Services)")

    add_subheading("3.6.5 ASYNCHRONOUS API MODULE (app/routes/api.py)")
    add_bullet("Auto-Save API (`/api/save-answer`): Receives asynchronous JSON payloads containing candidate answer selections and updates SQLite without page reloading.")
    add_bullet("Integrity Event API (`/api/integrity-event`): Records candidate window blur or tab switch incidents, logging timestamps and reducing the attempt's integrity score.")

    add_subheading("3.6.6 AUTOMATED SCORING SERVICE (app/services/scoring.py)")
    add_p("Executes server-side grading algorithms upon examination submission:")
    add_bullet("Single-Choice / True-False: Checks exact match between candidate response and answer key.")
    add_bullet("Multi-Choice: Computes proportional partial credit for correct subsets while penalizing incorrect selections.")
    add_bullet("Short Answer: Performs normalized, case-insensitive string matching against accepted answer keys.")

    add_subheading("3.6.7 VECTOR & AI SERVICES")
    add_bullet("Vector Service: ChromaDB embedding search for semantic question retrieval.")
    add_bullet("AI Feedback Service: Generates personalized candidate strength and weakness summaries.")

    # =========================================================================
    # PAGE 41: DIVIDER - TESTING
    # =========================================================================
    doc.add_page_break()
    add_divider_page("TESTING")

    # =========================================================================
    # PAGE 42: 4. TESTING AND IMPLEMENTATION
    # =========================================================================
    doc.add_page_break()
    add_page_title("4. TESTING AND IMPLEMENTATION")
    add_subheading("4.1 TEST CASES EXAMPLE")
    add_p("Testing is an essential phase in software engineering. It verifies whether the developed system functions correctly according to specified academic and technical requirements. Systematic testing identifies programming errors, unhandled edge cases, security vulnerabilities, and interface defects before production deployment.")

    add_subheading("Objectives of Testing")
    add_p("The primary objectives of testing the Online Examination System are:")
    add_bullet("1. To verify that user authentication and role-based access control operate securely.")
    add_bullet("2. To verify that examination configuration and question creation save accurate records.")
    add_bullet("3. To confirm that the server-authoritative timer prevents client-side clock tampering.")
    add_bullet("4. To verify that the asynchronous auto-save mechanism saves student answers in real time.")
    add_bullet("5. To ensure that student answers persist across page reloads and temporary disconnections.")
    add_bullet("6. To verify the accuracy of the automated grading algorithm and partial credit calculations.")
    add_bullet("7. To confirm that proctoring hooks log window blur and tab switch infractions accurately.")
    add_bullet("8. To verify that passing examinees receive verifiable digital certificates.")
    add_bullet("9. To validate the handling of invalid inputs and boundary conditions.")
    add_bullet("10. To ensure overall platform stability, responsiveness, and usability.")

    # =========================================================================
    # PAGE 43: 4.1 TEST CASES EXAMPLE (Testing Types)
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.1 TEST CASES EXAMPLE (Testing Types)")
    add_p("The application was subjected to comprehensive testing across multiple levels:")

    add_subheading("1. Unit Testing")
    add_p("Unit testing focuses on individual functions and classes in isolation. Unit tests were executed using Pytest to verify:")
    add_bullet("Password hashing and verification in the `User` model.")
    add_bullet("Remaining seconds calculation in the `Attempt` model.")
    add_bullet("Exact match and partial credit scoring formulas in `scoring.py`.")

    add_subheading("2. Integration Testing")
    add_p("Integration testing validates communication between different software modules:")
    add_bullet("Verifying that submitting an exam triggers the scoring service, updates the attempt status, and recalculates the student's dashboard records.")
    add_bullet("Ensuring that asynchronous auto-save requests update the `answers` table and reflect on the examination template.")

    add_subheading("3. System & Validation Testing")
    add_p("System testing evaluates the complete integrated web application from the perspective of students, instructors, and executive administrators:")
    add_bullet("Functional Testing: Confirms that every button, navigation link, and form functions as intended.")
    add_bullet("Validation Testing: Ensures the system rejects invalid inputs (e.g., duplicate emails, negative time limits).")

    # =========================================================================
    # PAGE 44: 4.1 TEST CASES EXAMPLE (Table Part 1: TC01 to TC07)
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.1 TEST CASES EXAMPLE (Execution Table Part 1)")
    add_p("The following structured test matrix details test cases executed for Authentication and Administrative functions:")

    tc_h1 = ["Test Case", "Module", "Test Description & Input", "Expected Result", "Status"]
    tc_r1 = [
        ["TC01", "Auth", "Admin enters valid email & password", "Admin dashboard opens successfully", "Pass"],
        ["TC02", "Auth", "Admin enters incorrect password", "Access denied; error message displayed", "Pass"],
        ["TC03", "Auth", "Register user with existing email", "Registration rejected; duplicate email alert", "Pass"],
        ["TC04", "Admin", "Create new exam with title & duration", "Exam created and saved in draft state", "Pass"],
        ["TC05", "Admin", "Add single-choice question with 4 options", "Question stored with correct option key", "Pass"],
        ["TC06", "Admin", "Add multi-choice question with 2 answers", "Options & answers saved as JSON arrays", "Pass"],
        ["TC07", "Admin", "Publish exam with complete question set", "Exam status toggled; visible to students", "Pass"]
    ]
    add_table_data(tc_h1, tc_r1, col_widths=[0.9, 0.9, 2.5, 1.9, 0.6])

    # =========================================================================
    # PAGE 45: 4.1 TEST CASES EXAMPLE (Table Part 2: TC08 to TC15)
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.1 TEST CASES EXAMPLE (Execution Table Part 2)")
    add_p("The following test matrix details test cases executed for Student Examination, Proctoring, and Scoring functions:")

    tc_h2 = ["Test Case", "Module", "Test Description & Input", "Expected Result", "Status"]
    tc_r2 = [
        ["TC08", "Student", "Student initiates published exam", "Attempt row created; backend timer set", "Pass"],
        ["TC09", "Student", "Student selects answer option", "Asynchronous API auto-saves response", "Pass"],
        ["TC10", "Student", "Browser page refreshed during exam", "Page reloads with saved answers intact", "Pass"],
        ["TC11", "Security", "Client alters device system clock", "Server timer ignores client; keeps true time", "Pass"],
        ["TC12", "Proctor", "Examinee switches browser window", "Event hook logs 'tab_switch' in audit table", "Pass"],
        ["TC13", "Scoring", "Student submits exam before expiry", "Auto-grading executes; score finalized", "Pass"],
        ["TC14", "Scoring", "Server timer reaches zero", "Server auto-submits & finalizes attempt", "Pass"],
        ["TC15", "Student", "Score >= pass_mark threshold", "Attempt marked passed; certificate issued", "Pass"]
    ]
    add_table_data(tc_h2, tc_r2, col_widths=[0.9, 0.9, 2.5, 1.9, 0.6])

    # =========================================================================
    # PAGE 46: 4.1 TEST CASES EXAMPLE (Testing Results & QA)
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.1 TEST CASES EXAMPLE (Testing Results & QA)")
    add_p("The testing process verified the major functions of the Online Examination System using Python. All 15 core test cases passed without defect, confirming that:")
    add_bullet("User authentication and role-based permissions operate reliably.")
    add_bullet("Exam configurations and question authoring forms maintain data consistency.")
    add_bullet("The server-authoritative timer strictly prevents duration extension.")
    add_bullet("Asynchronous auto-save ensures zero answer loss during browser interruptions.")
    add_bullet("Automated scoring accurately calculates exact matches and proportional partial credit.")
    add_bullet("Proctoring hooks reliably record window blur and tab switch infractions.")

    add_subheading("Testing Conclusion")
    add_p("Testing is essential for ensuring the quality, security, and stability of educational software. The Online Examination System has been tested module-by-module and as a complete end-to-end platform. The test results confirm that the system meets all functional requirements and can be deployed reliably for academic examinations.")

    # =========================================================================
    # PAGE 47: DIVIDER - IMPLEMENTATION
    # =========================================================================
    doc.add_page_break()
    add_divider_page("IMPLEMENTATION")

    # =========================================================================
    # PAGE 48: 4.2 IMPLEMENTATION
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.2 IMPLEMENTATION")
    add_p("Implementation is the stage where the designed system specifications and architectural plans are converted into a working software application. The Online Examination System using Python is implemented as a modular web application.")
    add_p("The implementation process followed a structured five-stage engineering workflow:")

    add_subheading("Stage 1: Environment & Dependency Setup")
    add_p("A virtual environment was configured to isolate project dependencies. Core libraries (Flask, Flask-SQLAlchemy, Flask-Login, Werkzeug, ChromaDB, Requests) were installed via `requirements.txt`.")

    add_subheading("Stage 2: Database Schema Creation & Seeding")
    add_p("SQLAlchemy ORM models were executed to construct database tables in SQLite (`data/oes.db`). Initial administrator, head, and student credentials were generated via `seed.py`.")

    add_subheading("Stage 3: Blueprint Routing & Template Binding")
    add_p("Modular route controllers (`auth`, `student`, `admin`, `head`, `api`) were registered on the application instance and bound to responsive Jinja2 HTML5 templates.")

    add_subheading("Stage 4: Client-Side Engine Integration")
    add_p("Client-side JavaScript (`exam.js`) was connected to handle real-time timer countdowns, auto-save API calls, and visibility change listeners.")

    add_subheading("Stage 5: Security Hardening")
    add_p("Password hashing, secure session cookies, and route permission decorators were verified to prevent unauthorized access.")

    # =========================================================================
    # PAGE 49: 4.2 IMPLEMENTATION (Architecture Flow)
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.2 IMPLEMENTATION (Architecture Diagram & Flow)")
    add_p("The system follows a tiered web application architecture where clients interact with the Flask application, which communicates with the SQLite database and specialized services.")

    arch_text = """+-----------------------------------------------------------------------------+
|                            USER ACTORS / ROLES                              |
|   [ Administrator ]              [ Student ]             [ Department Head ]|
|   - Manage Exams                 - View Exams            - Manage Users     |
|   - Manage Question Bank         - Take Timed Exams      - View Proctor Log |
|   - View Class Analytics         - View Scorecard & Cert - Institutional KPI|
+-------------------------------------+---------------------------------------+
                                      | HTTP / HTTPS Requests
                                      v
+-----------------------------------------------------------------------------+
|                            WEB USER INTERFACE                               |
|   - Responsive HTML5 Views, Glassmorphism CSS3 Design System                |
|   - Client-Side Exam Engine: Countdown Timer & Auto-Save JavaScript         |
+-------------------------------------+---------------------------------------+
                                      | WSGI Dispatching
                                      v
+-----------------------------------------------------------------------------+
|                         FLASK APPLICATION BACKEND                           |
|   - Auth Controller       - Student Controller       - Admin Controller     |
|   - Head Controller       - Auto-Save REST API       - Server Timer Engine  |
+-------------------+-----------------+-----------------------+---------------+
                    |                 |                       |
                    v                 v                       v
+-----------------------+ +-----------------------+ +-------------------------+
|   SQLITE DATABASE     | |   SCORING SERVICE     | | PROCTORING AUDIT ENGINE |
| - Users, Exams        | | - Exact Matching      | | - Window Blur Monitor   |
| - Questions, Attempts | | - Partial Credit      | | - Tab Switch Logger     |
| - Answers, Audit Logs | | - Grade Finalization  | | - Integrity Rating      |
+-----------------------+ +-----------------------+ +-------------------------+"""
    p_arch = doc.add_paragraph()
    p_arch.paragraph_format.space_before = Pt(8)
    p_arch.paragraph_format.space_after = Pt(8)
    p_arch.paragraph_format.line_spacing = 1.05
    r_ar = p_arch.add_run(arch_text)
    r_ar.font.name = 'Courier New'
    r_ar.font.size = Pt(8.5)

    # =========================================================================
    # PAGE 50: 4.3 DEPLOYMENT OPTIONS
    # =========================================================================
    doc.add_page_break()
    add_page_title("4.3 DEPLOYMENT OPTIONS")
    add_p("The Online Examination System is architected to accommodate varied institutional deployment requirements through three primary hosting configurations:")

    add_subheading("1. Campus Local Area Network (LAN) / On-Premise Deployment")
    add_bullet("Target Environment: Departmental testing centers and closed campus computer laboratories.")
    add_bullet("Architecture: Deployed on a campus server running Windows Server or Ubuntu Linux.")
    add_bullet("Web Server: Utilizes the production-grade Waitress WSGI server on Windows (`waitress-serve --port=5000 run:app`) or Gunicorn on Linux.")
    add_bullet("Data Storage: Local SQLite database with automated daily backups; requires zero external internet connection for core testing.")
    add_bullet("Advantages: Complete data sovereignty, ultra-low network latency, and complete immunity to external internet outages.")

    add_subheading("2. Containerized Cloud Deployment (Docker & Nginx)")
    add_bullet("Target Environment: Multi-campus universities conducting large-scale concurrent examinations.")
    add_bullet("Architecture: Packaged as a lightweight Docker container image with Python 3.12 runtime.")
    add_bullet("Web Server: Nginx reverse proxy managing SSL/TLS encryption and static caching, forwarding requests to Gunicorn workers.")
    add_bullet("Data Storage: SQLite or external PostgreSQL container by updating the database URI in `config.py`.")
    add_bullet("Advantages: Predictable execution, horizontal scalability, and high reliability.")

    add_subheading("3. Managed Cloud Platform-as-a-Service (PaaS)")
    add_bullet("Target Environment: Institutions seeking zero hardware maintenance overhead.")
    add_bullet("Platforms: Direct Git-based deployment on Render, Railway, or PythonAnywhere.")
    add_bullet("Advantages: Automated SSL certificates, zero server patching, and continuous delivery.")

    # =========================================================================
    # PAGES 51-62: APPENDICES - A) SAMPLE CODING (12 Pages of Real Code)
    # =========================================================================
    # Page 51: app/__init__.py
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n1. Flask Application Factory (app/__init__.py)", """from flask import Flask
from config import Config
from app.models.schema import db
from flask_login import LoginManager

login_manager = LoginManager()
login_manager.login_view = 'auth.login'
login_manager.login_message_category = 'warning'

def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Initialize extensions
    db.init_app(app)
    login_manager.init_app(app)

    # User loader callback for Flask-Login
    from app.models.schema import User
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    # Register Blueprints
    from app.routes.auth import auth_bp
    from app.routes.student import student_bp
    from app.routes.admin import admin_bp
    from app.routes.head import head_bp
    from app.routes.api import api_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(student_bp)
    app.register_blueprint(admin_bp, url_prefix='/admin')
    app.register_blueprint(head_bp, url_prefix='/head')
    app.register_blueprint(api_bp, url_prefix='/api')

    with app.app_context():
        db.create_all()

    return app""")

    # Page 52: Models: User & Exam
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n2. User & Exam Models (app/models/schema.py)", """from datetime import datetime, timezone
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

class User(UserMixin, db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='student')
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def is_admin(self):
        return self.role in ['admin', 'head']

    def is_head(self):
        return self.role == 'head'

class Exam(db.Model):
    __tablename__ = 'exams'
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text, nullable=True)
    duration_minutes = db.Column(db.Integer, nullable=False, default=30)
    pass_mark = db.Column(db.Float, nullable=False, default=50.0)
    is_published = db.Column(db.Boolean, default=False)
    created_by_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    questions = db.relationship('Question', backref='exam', cascade="all, delete-orphan")
    attempts = db.relationship('Attempt', backref='exam', cascade="all, delete-orphan")

    def total_points(self):
        return sum(q.points for q in self.questions)""")

    # Page 53: Models: Question & JSON Handling
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n3. Question Model with JSON Serializers (app/models/schema.py)", """import json

class Question(db.Model):
    __tablename__ = 'questions'
    id = db.Column(db.Integer, primary_key=True)
    exam_id = db.Column(db.Integer, db.ForeignKey('exams.id'), nullable=True)
    text = db.Column(db.Text, nullable=False)
    question_type = db.Column(db.String(30), nullable=False, default='single_choice')
    options_json = db.Column(db.Text, nullable=True)
    correct_answers_json = db.Column(db.Text, nullable=False)
    points = db.Column(db.Float, nullable=False, default=1.0)
    explanation = db.Column(db.Text, nullable=True)
    topic = db.Column(db.String(100), nullable=False, default='General')
    difficulty = db.Column(db.String(20), nullable=False, default='medium')
    created_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    @property
    def options(self):
        if self.options_json:
            try:
                return json.loads(self.options_json)
            except Exception:
                return []
        return []

    @options.setter
    def options(self, val):
        self.options_json = json.dumps(val)

    @property
    def correct_answers(self):
        if self.correct_answers_json:
            try:
                return json.loads(self.correct_answers_json)
            except Exception:
                return []
        return []

    @correct_answers.setter
    def correct_answers(self, val):
        self.correct_answers_json = json.dumps(val)""")

    # Page 54: Models: Attempt & Server-Authoritative Timer
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n4. Attempt Model & Server Timer Engine (app/models/schema.py)", """class Attempt(db.Model):
    __tablename__ = 'attempts'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    exam_id = db.Column(db.Integer, db.ForeignKey('exams.id'), nullable=False)
    start_time = db.Column(db.DateTime, nullable=False, default=lambda: datetime.now(timezone.utc))
    end_time = db.Column(db.DateTime, nullable=False) # Server authoritative target finish time
    submitted_at = db.Column(db.DateTime, nullable=True)
    status = db.Column(db.String(30), nullable=False, default='in_progress')
    score = db.Column(db.Float, nullable=True, default=0.0)
    max_score = db.Column(db.Float, nullable=False, default=0.0)
    passed = db.Column(db.Boolean, nullable=True, default=False)
    integrity_score = db.Column(db.Integer, default=100)
    certificate_id = db.Column(db.String(64), nullable=True, unique=True)

    answers = db.relationship('Answer', backref='attempt', cascade="all, delete-orphan")
    integrity_events = db.relationship('IntegrityEvent', backref='attempt', cascade="all, delete-orphan")

    def remaining_seconds(self):
        if self.status != 'in_progress':
            return 0
        now = datetime.now(timezone.utc)
        end = self.end_time
        if end.tzinfo is None:
            end = end.replace(tzinfo=timezone.utc)
        rem = (end - now).total_seconds()
        return max(0, int(rem))""")

    # Page 55: Models: Answer & IntegrityEvent
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n5. Answer & IntegrityEvent Models (app/models/schema.py)", """class Answer(db.Model):
    __tablename__ = 'answers'
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempts.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('questions.id'), nullable=False)
    student_response_json = db.Column(db.Text, nullable=True)
    is_correct = db.Column(db.Boolean, nullable=True)
    points_awarded = db.Column(db.Float, default=0.0)
    saved_at = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))

    question = db.relationship('Question', lazy=True)

    @property
    def student_response(self):
        if self.student_response_json:
            try:
                return json.loads(self.student_response_json)
            except Exception:
                return self.student_response_json
        return None

    @student_response.setter
    def student_response(self, val):
        self.student_response_json = json.dumps(val)

class IntegrityEvent(db.Model):
    __tablename__ = 'integrity_events'
    id = db.Column(db.Integer, primary_key=True)
    attempt_id = db.Column(db.Integer, db.ForeignKey('attempts.id'), nullable=False)
    event_type = db.Column(db.String(50), nullable=False) # 'tab_switch', 'window_blur', 'copy_attempt'
    details = db.Column(db.Text, nullable=True)
    timestamp = db.Column(db.DateTime, default=lambda: datetime.now(timezone.utc))""")

    # Page 56: Auth Controller
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n6. Authentication Controller (app/routes/auth.py)", """from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from app.models.schema import db, User

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('admin.dashboard' if current_user.is_admin() else 'student.dashboard'))
    if request.method == 'POST':
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        user = User.query.filter_by(email=email).first()
        if user and user.check_password(password):
            login_user(user)
            flash(f'Welcome back, {user.name}!', 'success')
            return redirect(url_for('admin.dashboard' if user.is_admin() else 'student.dashboard'))
        flash('Invalid institutional email or password.', 'danger')
    return render_template('auth/login.html')

@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        if User.query.filter_by(email=email).first():
            flash('Email already registered in system.', 'warning')
            return redirect(url_for('auth.register'))
        user = User(name=name, email=email, role='student')
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        flash('Registration successful! Please log in.', 'success')
        return redirect(url_for('auth.login'))
    return render_template('auth/register.html')""")

    # Page 57: Student Controller - Taking Exam
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n7. Student Exam Taking Controller (app/routes/student.py)", """from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_required, current_user
from datetime import datetime, timedelta, timezone
from app.models.schema import db, Exam, Attempt, Answer, Question

student_bp = Blueprint('student', __name__)

@student_bp.route('/exam/<int:exam_id>/take')
@login_required
def take_exam(exam_id):
    exam = Exam.query.get_or_404(exam_id)
    attempt = Attempt.query.filter_by(user_id=current_user.id, exam_id=exam_id, status='in_progress').first()
    
    if not attempt:
        start_time = datetime.now(timezone.utc)
        end_time = start_time + timedelta(minutes=exam.duration_minutes)
        attempt = Attempt(
            user_id=current_user.id,
            exam_id=exam_id,
            start_time=start_time,
            end_time=end_time,
            status='in_progress',
            max_score=exam.total_points()
        )
        db.session.add(attempt)
        db.session.commit()
        
        # Pre-seed blank answer entries for all questions
        for q in exam.questions:
            ans = Answer(attempt_id=attempt.id, question_id=q.id)
            db.session.add(ans)
        db.session.commit()
        
    remaining = attempt.remaining_seconds()
    if remaining <= 0:
        return redirect(url_for('student.submit_exam', attempt_id=attempt.id))
        
    return render_template('student/take_exam.html', exam=exam, attempt=attempt, remaining_seconds=remaining)""")

    # Page 58: Student Controller - Exam Submission & Scoring Trigger
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n8. Exam Submission & Result Finalization (app/routes/student.py)", """from app.services.scoring import score_attempt
import uuid

@student_bp.route('/attempt/<int:attempt_id>/submit', methods=['POST', 'GET'])
@login_required
def submit_exam(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    if attempt.user_id != current_user.id and not current_user.is_admin():
        flash('Unauthorized access to attempt.', 'danger')
        return redirect(url_for('student.dashboard'))
        
    if attempt.status == 'in_progress':
        attempt.submitted_at = datetime.now(timezone.utc)
        attempt.status = 'submitted'
        db.session.commit()
        
        # Execute automated grading algorithm
        score_attempt(attempt)
        
        # Generate verifiable certificate ID if student passed
        if attempt.passed and not attempt.certificate_id:
            attempt.certificate_id = f"CERT-OES-{uuid.uuid4().hex[:8].upper()}"
            db.session.commit()
            
        flash('Examination submitted and graded successfully!', 'success')
        
    return redirect(url_for('student.view_result', attempt_id=attempt.id))

@student_bp.route('/attempt/<int:attempt_id>/result')
@login_required
def view_result(attempt_id):
    attempt = Attempt.query.get_or_404(attempt_id)
    exam = attempt.exam
    answers = Answer.query.filter_by(attempt_id=attempt.id).all()
    score_pct = (attempt.score / attempt.max_score * 100) if attempt.max_score > 0 else 0
    return render_template('student/result.html', attempt=attempt, exam=exam, answers=answers, score_percentage=score_pct)""")

    # Page 59: Admin Controller - Exam & Question Management
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n9. Exam & Question Authoring Controller (app/routes/admin.py)", """from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from app.models.schema import db, Exam, Question

admin_bp = Blueprint('admin', __name__)

@admin_bp.route('/exam/new', methods=['GET', 'POST'])
@login_required
def create_exam():
    if not current_user.is_admin():
        flash('Admin privileges required.', 'danger')
        return redirect(url_for('student.dashboard'))
    if request.method == 'POST':
        title = request.form.get('title')
        desc = request.form.get('description')
        duration = int(request.form.get('duration_minutes', 30))
        pass_mark = float(request.form.get('pass_mark', 50.0))
        exam = Exam(title=title, description=desc, duration_minutes=duration, pass_mark=pass_mark, created_by_id=current_user.id)
        db.session.add(exam)
        db.session.commit()
        flash('Examination created successfully! You can now add questions.', 'success')
        return redirect(url_for('admin.manage_questions', exam_id=exam.id))
    return render_template('admin/create_exam.html')

@admin_bp.route('/exam/<int:exam_id>/publish')
@login_required
def publish_exam(exam_id):
    exam = Exam.query.get_or_404(exam_id)
    if len(exam.questions) == 0:
        flash('Cannot publish an exam with zero questions.', 'warning')
    else:
        exam.is_published = True
        db.session.commit()
        flash(f'Examination "{exam.title}" is now published and accessible to students.', 'success')
    return redirect(url_for('admin.exams'))""")

    # Page 60: Automated Scoring Algorithm
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n10. Automated Scoring & Partial Credit Engine (app/services/scoring.py)", """from app.models.schema import db

def score_attempt(attempt):
    total_score = 0.0
    for answer in attempt.answers:
        q = answer.question
        if not q:
            continue
        # Single Choice & True/False Evaluation
        if q.question_type in ['single_choice', 'true_false']:
            correct = q.correct_answers[0] if q.correct_answers else None
            student_ans = answer.student_response
            is_correct = (str(student_ans).strip().lower() == str(correct).strip().lower())
            answer.is_correct = is_correct
            answer.points_awarded = q.points if is_correct else 0.0
            
        # Multiple Choice with Proportional Partial Credit
        elif q.question_type == 'multi_choice':
            student_set = set(answer.student_response or [])
            correct_set = set(q.correct_answers)
            if student_set == correct_set:
                answer.is_correct = True
                answer.points_awarded = q.points
            elif student_set.issubset(correct_set) and len(student_set) > 0:
                answer.is_correct = False
                fraction = len(student_set) / len(correct_set)
                answer.points_awarded = round(q.points * fraction, 2)
            else:
                answer.is_correct = False
                answer.points_awarded = 0.0
                
        # Short Answer Evaluation
        elif q.question_type == 'short_answer':
            resp = str(answer.student_response or '').strip().lower()
            correct_keys = [str(k).strip().lower() for k in q.correct_answers]
            is_correct = resp in correct_keys
            answer.is_correct = is_correct
            answer.points_awarded = q.points if is_correct else 0.0
            
        total_score += answer.points_awarded
        
    attempt.score = total_score
    attempt.passed = (attempt.score / attempt.max_score * 100 >= attempt.exam.pass_mark) if attempt.max_score > 0 else False
    attempt.status = 'graded'
    db.session.commit()
    return attempt""")

    # Page 61: API & Proctoring Endpoints
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n11. Asynchronous Auto-Save & Proctoring API (app/routes/api.py)", """from flask import Blueprint, request, jsonify
from app.models.schema import db, Attempt, Answer, IntegrityEvent
import json

api_bp = Blueprint('api', __name__)

@api_bp.route('/save-answer', methods=['POST'])
def save_answer():
    data = request.get_json() or {}
    attempt_id = data.get('attempt_id')
    question_id = data.get('question_id')
    response_val = data.get('response')
    
    attempt = Attempt.query.get(attempt_id)
    if not attempt or attempt.status != 'in_progress':
        return jsonify({'error': 'Attempt closed or invalid'}), 400
        
    ans = Answer.query.filter_by(attempt_id=attempt_id, question_id=question_id).first()
    if not ans:
        ans = Answer(attempt_id=attempt_id, question_id=question_id)
        db.session.add(ans)
        
    ans.student_response = response_val
    db.session.commit()
    return jsonify({'success': True, 'message': 'Response saved asynchronously'})

@api_bp.route('/integrity-event', methods=['POST'])
def log_integrity_event():
    data = request.get_json() or {}
    attempt_id = data.get('attempt_id')
    event_type = data.get('event_type')
    details = data.get('details')
    
    attempt = Attempt.query.get(attempt_id)
    if attempt and attempt.status == 'in_progress':
        event = IntegrityEvent(attempt_id=attempt_id, event_type=event_type, details=details)
        db.session.add(event)
        # Deduct 5 integrity points per infraction
        attempt.integrity_score = max(0, (attempt.integrity_score or 100) - 5)
        db.session.commit()
        return jsonify({'success': True, 'integrity_score': attempt.integrity_score})
    return jsonify({'error': 'Invalid event'}), 400""")

    # Page 62: Client-Side Engine
    doc.add_page_break()
    add_code_page("APPENDICES - A) SAMPLE CODING\n12. Client-Side Proctoring & Timer Engine (app/static/js/exam.js)", """// Real-Time Server-Synchronized Timer Countdown
let timerInterval = null;
function startTimer(durationSeconds) {
    let remaining = durationSeconds;
    const timerDisplay = document.getElementById('exam-timer');
    
    timerInterval = setInterval(() => {
        remaining--;
        if (remaining <= 0) {
            clearInterval(timerInterval);
            timerDisplay.innerText = '00:00';
            alert('Time expired! Your examination is being automatically submitted.');
            document.getElementById('exam-form').submit();
            return;
        }
        const m = Math.floor(remaining / 60);
        const s = remaining % 60;
        timerDisplay.innerText = `${m < 10 ? '0' : ''}${m}:${s < 10 ? '0' : ''}${s}`;
    }, 1000);
}

// Proctoring Visibility Change Listener (Tab Switch Detection)
document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
        fetch('/api/integrity-event', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                attempt_id: ATTEMPT_ID,
                event_type: 'tab_switch',
                details: 'Candidate navigated away from active test window.'
            })
        }).then(res => res.json()).then(data => {
            console.warn('Integrity penalty logged. New rating:', data.integrity_score);
        });
    }
});

// Auto-Save Answer Selection
function autoSave(questionId, value) {
    const badge = document.getElementById('save-badge');
    badge.innerText = 'Saving...';
    fetch('/api/save-answer', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ attempt_id: ATTEMPT_ID, question_id: questionId, response: value })
    }).then(() => {
        badge.innerText = 'All responses saved';
    });
}""")

    # =========================================================================
    # PAGES 63-65: APPENDICES - B) SAMPLE INPUT (3 Pages)
    # =========================================================================
    # Page 63: User Authentication Inputs
    doc.add_page_break()
    add_page_title("APPENDICES - B) SAMPLE INPUT\n1. User Authentication & Registration Screens")
    add_p("This section illustrates the input structures, parameters, and form layouts for user authentication:")
    
    add_subheading("A. Student / Admin Login Form Input")
    add_p("The login interface accepts candidate and faculty credentials:")
    add_bullet("URL Route: `/login` (POST)")
    add_bullet("Parameter 1: `email` = 'logeshwari@college.edu' (User ID)")
    add_bullet("Parameter 2: `password` = '••••••••••••' (Encrypted text)")
    add_bullet("Action: Verifies user identity, initializes session cookie, redirects to user dashboard.")

    add_subheading("B. Student Registration Form Input")
    add_p("The self-registration interface registers candidates for academic testing:")
    add_bullet("URL Route: `/register` (POST)")
    add_bullet("Parameter 1: `name` = 'M. Logeshwari'")
    add_bullet("Parameter 2: `email` = 'logeshwari@college.edu'")
    add_bullet("Parameter 3: `password` = 'CandidatePass#2026'")
    add_bullet("Parameter 4: `role` = 'student' (Default assignment)")
    add_bullet("Validation: Confirms RFC email format and enforces non-duplicate email constraint.")

    # Page 64: Exam & Question Creation Inputs
    doc.add_page_break()
    add_page_title("APPENDICES - B) SAMPLE INPUT\n2. Exam & Question Authoring Inputs")
    add_p("This section illustrates the structured form payloads submitted by faculty administrators:")

    add_subheading("A. Examination Definition Payload")
    add_bullet("URL Route: `/admin/exam/new` (POST)")
    add_bullet("Exam Title: 'Python Programming and Data Structures Assessment'")
    add_bullet("Description: 'Comprehensive test evaluating syntax, control flow, lists, and OOP concepts.'")
    add_bullet("Duration: '45' (Minutes allowed)")
    add_bullet("Pass Mark: '60.0' (Passing threshold percentage)")

    add_subheading("B. Question Authoring Payload (Multi-Choice)")
    add_bullet("URL Route: `/admin/exam/1/question/new` (POST)")
    add_bullet("Question Text: 'Which of the following built-in data types in Python are mutable?'")
    add_bullet("Question Type: 'multi_choice'")
    add_bullet("Options: '[\"List\", \"Tuple\", \"Dictionary\", \"String\"]'")
    add_bullet("Correct Answers: '[\"List\", \"Dictionary\"]'")
    add_bullet("Points: '2.0'")
    add_bullet("Topic: 'Python Collections'")
    add_bullet("Difficulty: 'medium'")
    add_bullet("Explanation: 'Lists and dictionaries can be modified in place; tuples and strings cannot.'")

    # Page 65: Active Exam & Auto-Save Inputs
    doc.add_page_break()
    add_page_title("APPENDICES - B) SAMPLE INPUT\n3. Live Examination Response & Auto-Save Payloads")
    add_p("This section illustrates the real-time input interactions transmitted by the student browser:")

    add_subheading("A. Asynchronous Answer Auto-Save Payload")
    add_bullet("URL Route: `/api/save-answer` (POST JSON)")
    add_bullet("Payload: `{\"attempt_id\": 102, \"question_id\": 5, \"response\": \"List\"}`")
    add_bullet("Trigger: Fired immediately when candidate clicks a radio button or checkbox.")
    add_bullet("Validation: Confirms active attempt session; writes response to SQLite database.")

    add_subheading("B. Client Proctoring Infraction Payload")
    add_bullet("URL Route: `/api/integrity-event` (POST JSON)")
    add_bullet("Payload: `{\"attempt_id\": 102, \"event_type\": \"tab_switch\", \"details\": \"Candidate unfocused active window.\"}`")
    add_bullet("Trigger: Triggered by the browser `visibilitychange` or `blur` event.")
    add_bullet("System Action: Appends record to `integrity_events` table and reduces integrity rating by 5%.")

    # =========================================================================
    # PAGES 66-71: APPENDICES - C) SAMPLE OUTPUT (6 Pages)
    # =========================================================================
    # Page 66: Admin Dashboard Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n1. Administrative Dashboard & Analytics")
    add_p("The Administrative Dashboard displays institution-wide assessment metrics and quick action controls:")

    out_1 = """+-----------------------------------------------------------------------------+
| ONLINE EXAMINATION SYSTEM - ADMINISTRATIVE DASHBOARD                        |
| Logged in: Admin (admin@oes.com)                     Session: Active        |
+-----------------------------------------------------------------------------+
| [ METRIC OVERVIEW CARDS ]                                                   |
|   TOTAL EXAMINATIONS : 12 Active Assessments                                |
|   QUESTION BANK      : 148 Verified Questions                               |
|   STUDENT ATTEMPTS   : 342 Recorded Submissions                             |
|   AVERAGE SCORE      : 74.8% Global Class Average                           |
+-----------------------------------------------------------------------------+
| [ ACTIVE EXAMINATIONS REPOSITORY ]                                          |
|   Title: Python Programming Fundamentals        Status: Published (45 min)  |
|   Questions: 25 items | Pass Mark: 60%          Attempts: 84 completed      |
|   Actions: [ View Questions ]  [ Edit Config ]  [ Performance Analytics ]   |
|                                                                             |
|   Title: Object-Oriented Software Design        Status: Published (60 min)  |
|   Questions: 30 items | Pass Mark: 50%          Attempts: 62 completed      |
|   Actions: [ View Questions ]  [ Edit Config ]  [ Performance Analytics ]   |
+-----------------------------------------------------------------------------+"""
    p_o1 = doc.add_paragraph()
    p_o1.paragraph_format.line_spacing = 1.05
    r_o1 = p_o1.add_run(out_1)
    r_o1.font.name = 'Courier New'
    r_o1.font.size = Pt(8.5)

    # Page 67: Student Dashboard Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n2. Student Portal & Assessment Ledger")
    add_p("The Student Dashboard presents available tests and past assessment records:")

    out_2 = """+-----------------------------------------------------------------------------+
| ONLINE EXAMINATION SYSTEM - STUDENT PORTAL                                  |
| Candidate: M. LOGESHWARI (3-B.Sc Computer Science)   Status: Enrolled       |
+-----------------------------------------------------------------------------+
| [ AVAILABLE ASSESSMENTS ]                                                   |
|   Exam: Python Programming and Data Structures Assessment                   |
|   Duration: 45 Minutes | Questions: 20 Items | Pass Mark: 60.0%             |
|   Status: READY FOR TESTING                                                 |
|   Action: [ >>> START EXAMINATION <<< ]                                     |
+-----------------------------------------------------------------------------+
| [ PAST ATTEMPTS & SCORE HISTORY ]                                           |
|   Date         Exam Title                     Score     Result    Action    |
|   2026-09-15   Database Management Systems    24/30     PASSED    [Review]  |
|   2026-09-18   Web Technologies & UI Design   27/30     PASSED    [Review]  |
|   2026-09-19   Python Programming Assessment  26/30     PASSED    [Cert]    |
+-----------------------------------------------------------------------------+"""
    p_o2 = doc.add_paragraph()
    p_o2.paragraph_format.line_spacing = 1.05
    r_o2 = p_o2.add_run(out_2)
    r_o2.font.name = 'Courier New'
    r_o2.font.size = Pt(8.5)

    # Page 68: Live Exam Interface Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n3. Live Examination Taking Interface")
    add_p("The live testing interface presents the real-time timer countdown, auto-save badge, and question palette:")

    out_3 = """+-----------------------------------------------------------------------------+
| Python Programming and Data Structures Assessment                           |
| Candidate: M. Logeshwari     [ Remaining Time: 38:42 ]   [ All Saved ]      |
+-----------------------------------------------------------------------------+
| Question 4 of 20                                               [ 2.0 Points ]
| Topic: Collections & Mutability                                             |
|                                                                             |
| Which of the following built-in data types in Python are mutable?           |
| (Select all that apply)                                                     |
|                                                                             |
|   [X] A) List                                                               |
|   [ ] B) Tuple                                                              |
|   [X] C) Dictionary                                                         |
|   [ ] D) String                                                             |
|                                                                             |
| [ < Previous ]                                                [ Next > ]    |
+-----------------------------------------------------------------------------+
| QUESTION PALETTE:                                                           |
| [ 1:SAVED ] [ 2:SAVED ] [ 3:SAVED ] [ 4:ACTIVE ] [ 5:EMPTY ] [ 6:EMPTY ] ... |
|                                                    [ SUBMIT FINAL EXAM ]    |
+-----------------------------------------------------------------------------+"""
    p_o3 = doc.add_paragraph()
    p_o3.paragraph_format.line_spacing = 1.05
    r_o3 = p_o3.add_run(out_3)
    r_o3.font.name = 'Courier New'
    r_o3.font.size = Pt(8.5)

    # Page 69: Scorecard Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n4. Examination Scorecard & Detailed Review")
    add_p("The candidate scorecard displays final grades, pass/fail status, integrity ratings, and explanations:")

    out_4 = """================================================================================
                    ONLINE EXAMINATION OFFICIAL SCORECARD
================================================================================
Candidate Name    : M. LOGESHWARI
Degree / Class    : 3-B.Sc Computer Science
Assessment Title  : Python Programming and Data Structures Assessment
Date Completed    : 2026-09-19 10:45:00 UTC
Session Duration  : 42 Minutes 18 Seconds
--------------------------------------------------------------------------------
FINAL SCORE       : 26.0 / 30.0 Points
SCORE PERCENTAGE  : 86.7% (Pass Mark: 60.0%)
RESULT STATUS     : PASSED [FIRST CLASS WITH DISTINCTION]
INTEGRITY RATING  : 100% [Zero suspicious infractions recorded]
CERTIFICATE NO    : CERT-OES-88A4DF
--------------------------------------------------------------------------------
DETAILED ITEM REVIEW:
  * Question 1: [CORRECT] Single Choice - Variable Scoping (Points: 1.0/1.0)
    Explanation: Python follows LEGB (Local, Enclosing, Global, Built-in) rule.
  * Question 4: [CORRECT] Multi-Choice - Mutable Collections (Points: 2.0/2.0)
    Explanation: Lists and dictionaries support in-place mutation.
================================================================================"""
    p_o4 = doc.add_paragraph()
    p_o4.paragraph_format.line_spacing = 1.05
    r_o4 = p_o4.add_run(out_4)
    r_o4.font.name = 'Courier New'
    r_o4.font.size = Pt(8.5)

    # Page 70: Digital Certificate Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n5. Verifiable Digital Certificate of Achievement")
    add_p("The platform automatically issues verifiable completion certificates for passing scores:")

    out_5 = """+------------------------------------------------------------------------------+
|               PACHAMUTHU COLLEGE OF ARTS AND SCIENCE FOR WOMEN               |
|                         DEPARTMENT OF COMPUTER SCIENCE                       |
|                                                                              |
|                    CERTIFICATE OF ACADEMIC ACHIEVEMENT                       |
|                                                                              |
|  This is to certify that                                                     |
|                               M. LOGESHWARI                                  |
|                         (3-B.Sc Computer Science)                            |
|                                                                              |
|  has successfully passed the institutional examination:                      |
|               ONLINE EXAMINATION SYSTEM - PYTHON PROGRAMMING                 |
|                                                                              |
|  Achieving a distinction score of 86.7% with an Integrity Rating of 100%.    |
|                                                                              |
|  Issue Date: September 19, 2026                 Certificate ID: CERT-OES-88 |
|                                                                              |
|  Ms. C. MOHANAPRIYA, M.Sc.                       Dr. J. THAVAMANI            |
|  Head of Department                              Principal                   |
+------------------------------------------------------------------------------+"""
    p_o5 = doc.add_paragraph()
    p_o5.paragraph_format.line_spacing = 1.05
    r_o5 = p_o5.add_run(out_5)
    r_o5.font.name = 'Courier New'
    r_o5.font.size = Pt(8.5)

    # Page 71: Executive Proctoring Audit Log Output
    doc.add_page_break()
    add_page_title("APPENDICES - C) SAMPLE OUTPUT\n6. Executive Proctoring & Integrity Audit Log")
    add_p("The Executive Audit Log displays real-time candidate integrity logs across all active sessions:")

    out_6 = """================================================================================
                    EXECUTIVE PROCTORING & INTEGRITY AUDIT LOG
================================================================================
Timestamp (UTC)      Attempt ID  Candidate Email      Event Type   Integrity
2026-09-19 10:12:04  ATT-101     student2@oes.com     tab_switch   95% (-5%)
2026-09-19 10:14:22  ATT-101     student2@oes.com     copy_attempt 90% (-5%)
2026-09-19 10:20:15  ATT-102     student3@oes.com     tab_switch   95% (-5%)
2026-09-19 10:30:11  ATT-104     logeshwari@oes.com   none (clean) 100%
================================================================================
AUDIT SUMMARY:
  * Total Concurrent Sessions Monitored : 42 Attempts
  * Flagged Infraction Incidents        : 3 Logged
  * Verified Clean Sessions             : 39 Attempts (92.8% Compliance)
================================================================================"""
    p_o6 = doc.add_paragraph()
    p_o6.paragraph_format.line_spacing = 1.05
    r_o6 = p_o6.add_run(out_6)
    r_o6.font.name = 'Courier New'
    r_o6.font.size = Pt(8.5)

    # =========================================================================
    # PAGE 72: DIVIDER - CONCLUSION
    # =========================================================================
    doc.add_page_break()
    add_divider_page("CONCLUSION")

    # =========================================================================
    # PAGE 73: 5. CONCLUSION (Project Summary & Deliverables)
    # =========================================================================
    doc.add_page_break()
    add_page_title("5. CONCLUSION")
    add_subheading("5.1 PROJECT SUMMARY")
    add_p("The Online Examination System using Python is a web-based application developed to simplify, automate, and safeguard institutional examinations and academic grading.")
    add_p("The system provides a centralized platform for administrators, faculty educators, and student candidates. Administrators can author multi-format questions, schedule assessments, enforce passing thresholds, and review class analytics. Students can securely log in to access scheduled tests, track remaining time, select answers with automatic background persistence, and access finalized scorecards and verifiable certificates.")
    add_p("The system eliminates manual paper distribution, saves printing costs, and prevents human clerical calculation errors. The use of an SQLite transactional database ensures fast, reliable access to student records and assessment archives.")
    add_p("The server-authoritative timer prevents client-side duration tampering, while the proctoring audit log tracks candidate window focus losses. Furthermore, objective questions are graded instantly, delivering immediate performance feedback.")
    add_p("Overall, the Online Examination System provides a simple, centralized, and efficient solution for modern educational institutions.")

    # =========================================================================
    # PAGE 74: 5. CONCLUSION (Future Enhancements & Final Statement)
    # =========================================================================
    doc.add_page_break()
    add_page_title("5.2 FUTURE ENHANCEMENTS & FINAL STATEMENT")
    add_p("The system can be further improved by incorporating the following future enhancements:")
    add_bullet("1. Webcam AI Proctoring – Computer vision models (OpenCV / MediaPipe) can be integrated to detect multiple faces or smartphone usage during examinations.")
    add_bullet("2. Coding Sandbox Assessment – Embedded code compilation sandboxes can be introduced for computer science programming assessments.")
    add_bullet("3. Mobile Application – Dedicated Android and iOS applications can be developed using Flutter for offline-first testing.")
    add_bullet("4. SMS and Email Alerts – Important assessment announcements, exam schedules, and grade publications can be transmitted via automated gateways.")
    add_bullet("5. Cloud Database Migration – SQLite can be upgraded to cloud-hosted PostgreSQL for massive nationwide testing scale.")

    add_subheading("5.3 FINAL STATEMENT")
    add_p("The developed system successfully demonstrates how Python web technologies and software engineering principles can be applied to transform academic testing. It provides a convenient, paperless, and secure method for managing examination workflows and establishes a resilient foundation for modern digital education.")

    # =========================================================================
    # PAGE 75: DIVIDER - BIBLIOGRAPHY
    # =========================================================================
    doc.add_page_break()
    add_divider_page("BIBLIOGRAPHY")

    # =========================================================================
    # PAGE 76: 6. BIBLIOGRAPHY (Complete References)
    # =========================================================================
    doc.add_page_break()
    add_page_title("6. BIBLIOGRAPHY")
    add_p("The following books, official documentation, and technical standards were referred to during the development of the Online Examination System using Python:")

    add_subheading("6.1 BOOKS")
    add_bullet("1. Roger S. Pressman, Bruce R. Maxim, Software Engineering: A Practitioner's Approach, 9th ed., McGraw-Hill Education, 2020.")
    add_bullet("2. Ian Sommerville, Software Engineering, 10th ed., Pearson Education, 2016.")
    add_bullet("3. Miguel Grinberg, Flask Web Development: Developing Web Applications with Python, 2nd ed., O'Reilly Media, 2018.")
    add_bullet("4. Abraham Silberschatz, Henry F. Korth, S. Sudarshan, Database System Concepts, 7th ed., McGraw-Hill Education, 2019.")
    add_bullet("5. Mark Lutz, Learning Python, 5th ed., O'Reilly Media, 2013.")

    add_subheading("6.2 ONLINE RESOURCES")
    add_bullet("1. Python Documentation – https://docs.python.org/3/")
    add_bullet("2. Flask Web Framework Documentation – https://flask.palletsprojects.com/")
    add_bullet("3. SQLite Database Engine Documentation – https://www.sqlite.org/docs.html")
    add_bullet("4. SQLAlchemy ORM Documentation – https://docs.sqlalchemy.org/")
    add_bullet("5. MDN Web Docs: HTML5, CSS3, and JavaScript – https://developer.mozilla.org/")

    add_subheading("6.3 DEVELOPMENT TOOLS & TECHNICAL STANDARDS")
    add_bullet("1. Visual Studio Code – Integrated development environment used for code authoring.")
    add_bullet("2. Google Chrome & Microsoft Edge – Browsers used for client-side DOM testing.")
    add_bullet("3. OWASP Foundation Web Application Security Standards – https://owasp.org/")

    # Save document
    output_filename = "Online_Examination_System_76_Pages_Report.docx"
    doc.save(output_filename)
    print(f"Successfully generated 76-page report: '{output_filename}'")
    # Also attempt to save to primary file if not locked by Word
    try:
        doc.save("Online_Examination_System_Project_Report.docx")
        print("Also updated 'Online_Examination_System_Project_Report.docx'")
    except Exception as e:
        print(f"Note: Could not overwrite Online_Examination_System_Project_Report.docx (likely open in Word): {e}")

if __name__ == '__main__':
    create_76_page_report()
