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

def build_complete_project_document():
    doc = docx.Document()

    # Set 1-inch margins
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Times New Roman'
    normal_font.size = Pt(12)
    normal_font.color.rgb = RGBColor(0, 0, 0)

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(30)
        p.paragraph_format.space_after = Pt(16)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_heading_3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.bold = True
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_p(text, bold_prefix=None, space_after=6, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(12)
            run_b.bold = True
            run_b.font.color.rgb = RGBColor(0, 0, 0)
        run_t = p.add_run(text)
        run_t.font.name = 'Times New Roman'
        run_t.font.size = Pt(12)
        run_t.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(12)
            run_b.bold = True
            run_b.font.color.rgb = RGBColor(0, 0, 0)
        run_t = p.add_run(text)
        run_t.font.name = 'Times New Roman'
        run_t.font.size = Pt(12)
        run_t.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_code_block(code_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(code_text)
        run.font.name = 'Courier New'
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(30, 30, 30)
        return p

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
                        run.font.size = Pt(10)

        if col_widths:
            for row in tbl.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)

        sp_p = doc.add_paragraph()
        sp_p.paragraph_format.space_before = Pt(0)
        sp_p.paragraph_format.space_after = Pt(6)

    # =========================================================================
    # PAGE 1: TITLE PAGE
    # =========================================================================
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(40)
    p1.paragraph_format.space_after = Pt(12)
    r1 = p1.add_run("Online Examination System using Python")
    r1.font.name = 'Times New Roman'
    r1.font.size = Pt(18)
    r1.bold = True

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(12)
    p2.paragraph_format.space_after = Pt(8)
    r2 = p2.add_run("A major project work submitted in partial\nfulfillment of the requirements for the degree of\n")
    r2.font.name = 'Times New Roman'
    r2.font.size = Pt(12)

    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(8)
    p3.paragraph_format.space_after = Pt(8)
    r3 = p3.add_run("3-B.Sc Computer Science\n(BACHELOR OF SCIENCE IN COMPUTER SCIENCE)")
    r3.font.name = 'Times New Roman'
    r3.font.size = Pt(13)
    r3.bold = True

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(12)
    p4.paragraph_format.space_after = Pt(16)
    r4 = p4.add_run("to the\nPeriyar University, Salem – 636011\n\nBy\n")
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
    # PAGE 2: CERTIFICATE / BONAFIDE
    # =========================================================================
    doc.add_page_break()
    add_title("CERTIFICATE")
    
    p_cert = doc.add_paragraph()
    p_cert.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cert.paragraph_format.space_before = Pt(16)
    p_cert.paragraph_format.space_after = Pt(16)
    r_c = p_cert.add_run("PACHAMUTHU COLLEGE OF ARTS AND SCIENCE FOR WOMEN\n(AFFILIATED TO PERIYAR UNIVERSITY)\nDHARMAPURI – 636701\n\nPROJECT WORK – SEPTEMBER 2026\n\nBONAFIDE WORK DONE")
    r_c.font.name = 'Times New Roman'
    r_c.font.size = Pt(12)
    r_c.bold = True

    add_p("This is to certify that the project entitled \"Online Examination System using Python\" is a bonafide record of work carried out by M.LOGESHWARI, Degree: 3-B.Sc Computer Science, in partial fulfillment of the requirements for the award of the degree of Bachelor of Science in Computer Science to Periyar University, Salem, during the academic year 2025–2026.")
    add_p("The project report has been approved as satisfying the academic requirements for the major project submission.")

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_sig.paragraph_format.space_before = Pt(40)
    p_sig.paragraph_format.space_after = Pt(20)
    r_sig = p_sig.add_run("INTERNAL GUIDE                                                HEAD OF THE DEPARTMENT\n(Ms. C. MOHANAPRIYA, M.Sc.)                           (Department of Computer Science)\n\n\nSUBMITTED FOR THE VIVA-VOCE EXAMINATION HELD ON: ____________________\n\n\nINTERNAL EXAMINER                                            EXTERNAL EXAMINER")
    r_sig.font.name = 'Times New Roman'
    r_sig.font.size = Pt(11)
    r_sig.bold = True

    # =========================================================================
    # PAGE 3: ACKNOWLEDGMENT
    # =========================================================================
    doc.add_page_break()
    add_title("ACKNOWLEDGMENT")
    add_p("At the outset, I offer my humble prayers to the Almighty and to my parents for giving me the strength, perseverance, and determination not only in the pursuit of this task but also in all walks of life.")
    add_p("I express my sincere thanks to our honourable Principal Dr. J. THAVAMANI, M.Com., M.Phil., Ph.D., for providing the necessary facilities and academic environment to carry out my project work successfully.")
    add_p("My wholehearted and profound gratitude to Ms. C. MOHANAPRIYA, M.Sc., Head of the Department of Computer Science, for her invaluable cooperation, continuous guidance, and encouragement given to me at every step throughout this project.")
    add_p("My wholehearted and proud gratitude goes to my project guide Ms. C. MOHANAPRIYA, M.Sc., for supervising me in choosing this project topic, verifying design artifacts, and enlightening my endeavor throughout this project work.")
    add_p("I thank all the respected faculty members of the Department of Computer Science for their academic support, constructive feedback, and technical guidance. Finally, my gratitude goes to my parents, classmates, and friends who have made direct and indirect contributions to make this project a grand success.")

    # =========================================================================
    # PAGE 4: INDEX PAGE (TABLE OF CONTENTS)
    # MUST ONLY CONTAIN THE HEADINGS SPECIFIED BY THE USER
    # =========================================================================
    doc.add_page_break()
    add_title("INDEX")

    index_headers = ["CHAPTER / SECTION", "TITLE", "PAGE NO"]
    index_rows = [
        ["1", "INTRODUCTION", "1"],
        ["", "  1.1 SYSTEM SPECIFICATION", "2"],
        ["", "    1.1.1 HARDWARE CONFIGURATION", "2"],
        ["", "    1.1.2 SOFTWARE CONFIGURATION", "3"],
        ["2", "SYSTEM STUDY", "5"],
        ["", "  2.1 EXISTING SYSTEM", "5"],
        ["", "    2.1.1 DESCRIPTION", "5"],
        ["", "    2.1.2 DRAWBACKS", "6"],
        ["", "  2.2 PROPOSED SYSTEM", "7"],
        ["", "    2.2.1 DESCRIPTION", "7"],
        ["", "    2.2.2 FEATURES", "8"],
        ["3", "SYSTEM DESIGN AND DEVELOPMENT", "10"],
        ["", "  3.1 FILE DESIGN", "10"],
        ["", "  3.2 INPUT DESIGN", "11"],
        ["", "  3.3 OUTPUT DESIGN", "13"],
        ["", "  3.4 DATABASE DESIGN", "15"],
        ["", "  3.5 SYSTEM DEVELOPMENT", "18"],
        ["", "  3.6 DESCRIPTION OF MODULES", "19"],
        ["4", "TESTING AND IMPLEMENTATION", "22"],
        ["", "  4.1 TEST CASES EXAMPLE", "22"],
        ["", "  4.2 IMPLEMENTATION", "24"],
        ["", "  4.3 DEPLOYMENT OPTIONS", "26"],
        ["5", "CONCLUSION", "28"],
        ["6", "BIBLIOGRAPHY", "29"],
        ["7", "APPENDICES", "30"],
        ["", "  A) SAMPLE CODING", "30"],
        ["", "  B) SAMPLE INPUT", "35"],
        ["", "  C) SAMPLE OUTPUT", "37"]
    ]
    add_table_data(index_headers, index_rows, col_widths=[1.5, 4.0, 1.0])

    # =========================================================================
    # CHAPTER 1: INTRODUCTION
    # =========================================================================
    doc.add_page_break()
    add_heading_1("1. INTRODUCTION")
    add_p("The Online Examination System using Python is a web-based academic assessment platform engineered to simplify, automate, and safeguard the examination process for educational institutions. In traditional educational setups, conducting examinations involves extensive logistical coordination, including printing question papers, scheduling physical invigilators, managing examination halls, collecting answer sheets, manually evaluating papers, and manually compiling score reports. This conventional process is not only time-consuming and expensive but is also prone to human calculation errors, delays in result publishing, and vulnerability to academic infractions.")
    add_p("With the modernization of academic workflows, educational institutions increasingly demand digital examination software that combines accessibility with rigorous examination integrity. The Online Examination System meets these demands by providing a secure, centralized web portal developed using the Python Flask framework and an SQLite relational database. The platform supports multiple administrative and academic tiers, providing dedicated features for Student examinees, Instructors/Administrators, and Executive Department Heads.")
    add_p("Key operational objectives accomplished by this project include:")
    add_bullet("Elimination of paper waste and manual logistics by converting traditional written exams into responsive digital assessments.", "Paperless Evaluation: ")
    add_bullet("Enforcement of backend server-authoritative timers to ensure students cannot tamper with test durations via client browser adjustments.", "Exam Integrity & Anti-Tamper: ")
    add_bullet("Implementation of asynchronous background auto-save endpoints that continuously persist student responses, ensuring zero data loss during network hiccups.", "Asynchronous Resilience: ")
    add_bullet("Support for diverse question types including single choice, multiple choice with partial credit, true/false, and short answer questions.", "Multi-Format Question Bank: ")
    add_bullet("Instant evaluation of objective questions upon submission, providing students with immediate performance scorecards and verifiable certificates.", "Immediate Feedback: ")
    add_bullet("Comprehensive executive proctoring and audit logging that tracks window focus losses, tab switches, and clipboard infractions in real time.", "Real-Time Proctor Audit: ")

    add_heading_2("1.1 SYSTEM SPECIFICATION")
    add_p("The system specification defines the essential hardware and software requirements necessary to successfully develop, execute, host, and access the Online Examination System. Adhering to these specifications guarantees predictable performance, secure database transactions, and a fluid user experience.")

    add_heading_3("1.1.1 HARDWARE CONFIGURATION")
    add_p("The hardware configuration specifies the minimum and recommended physical computing resources needed for the host server and client access nodes:")
    add_bullet("Intel Core i3 processor (2.0 GHz or higher) or AMD equivalent. A multi-core processor (Intel Core i5 or above) is recommended for high-volume concurrent testing.", "Server / Host Processor: ")
    add_bullet("4 GB RAM minimum; 8 GB RAM recommended for multi-user session management and concurrent database connections.", "System Memory (RAM): ")
    add_bullet("500 GB Hard Disk Drive (HDD) or 256 GB Solid State Drive (SSD) with at least 5 GB free disk storage for application files, logs, and database records.", "Storage Space: ")
    add_bullet("Standard Color LED/LCD Monitor supporting 1366x768 resolution or Full HD (1920x1080) for clear presentation of examination dashboards.", "Display: ")
    add_bullet("Standard 104-key USB/Wireless Keyboard and Optical Scroll Mouse for student inputs and administrative data entry.", "Input Peripherals: ")
    add_bullet("Broadband LAN or Wi-Fi connection (minimum 2 Mbps per client) for real-time exam communication and asynchronous response auto-saving.", "Network Connectivity: ")
    add_bullet("Standard desktop PC, laptop, or tablet with any standard web browser; no proprietary client hardware installation required.", "Client Workstation: ")

    add_heading_3("1.1.2 SOFTWARE CONFIGURATION")
    add_p("The software configuration encompasses the operating environment, development runtimes, libraries, and frameworks that comprise the application stack:")
    add_bullet("Microsoft Windows 10 / Windows 11 (64-bit), or Linux distributions (Ubuntu 20.04 LTS / Debian 11+).", "Operating System: ")
    add_bullet("Python 3.10 to 3.13 (High-level, interpreted language providing backend stability, clean syntax, and rich standard library support).", "Programming Language: ")
    add_bullet("Flask (Lightweight, robust Python WSGI micro-framework supporting routing, blueprints, templating, and session handling).", "Web Backend Framework: ")
    add_bullet("SQLite 3 (Serverless, transactional, self-contained relational database management system storing exams, questions, users, and audit records).", "Relational Database: ")
    add_bullet("HTML5 (Semantic markup for accessible forms and interactive exam question cards).", "Markup Language: ")
    add_bullet("CSS3 (Custom styling with modern glassmorphism design tokens, CSS variables, and responsive grid layouts).", "Styling Language: ")
    add_bullet("Vanilla JavaScript (ES6+ for asynchronous Fetch API calls, real-time timer countdowns, and browser event listeners).", "Client-Side Scripting: ")
    add_bullet("Flask-SQLAlchemy (ORM for database mapping), Flask-Login (session authentication), Werkzeug (password hashing security).", "Core Python Libraries: ")
    add_bullet("Google Chrome 110+, Microsoft Edge 110+, or Mozilla Firefox 110+ with active JavaScript runtime.", "Web Browser: ")
    add_bullet("Visual Studio Code (VS Code) with Python and Jinja extension packs for source code development and debugging.", "Integrated Development Environment (IDE): ")
    add_bullet("Flask Built-in Development WSGI Server for local testing; Waitress / Gunicorn for production deployments.", "Application Web Server: ")

    # =========================================================================
    # CHAPTER 2: SYSTEM STUDY
    # =========================================================================
    doc.add_page_break()
    add_heading_1("2. SYSTEM STUDY")
    add_p("System Study is an indispensable preliminary phase in software engineering. It involves conducting a rigorous examination of the existing operational methodologies, cataloging operational limitations and vulnerabilities, determining functional user requirements, and formulating a viable, technologically sound proposed system. This study ensures that the developed solution effectively solves actual real-world institutional assessment problems.")

    add_heading_2("2.1 EXISTING SYSTEM")
    add_heading_3("2.1.1 DESCRIPTION")
    add_p("In the existing conventional examination system, educational institutions primarily rely on manual paper-based examinations or rudimentary third-party survey forms (such as basic Google Forms). In the paper-based model, faculty members draft questions manually, submit them to an examination committee for duplication, and physically store question papers under security until exam day. Students assemble in designated examination halls where invigilators distribute physical papers and monitor time using classroom wall clocks. Completed paper scripts are physically collected, organized by registration number, and distributed to professors for manual marking. Scores are then transcribed into record registers or spreadsheets.")
    add_p("Where basic computerized forms are used, they typically operate without server-side validation. Timers in such forms run solely through client-side JavaScript intervals. Furthermore, basic online forms do not persist responses dynamically, requiring examinees to reach the final page and click a single submit button to save any of their work.")

    add_heading_3("2.1.2 DRAWBACKS")
    add_p("The existing examination methodology suffers from substantial structural, financial, and operational drawbacks:")
    add_bullet("Printing thousands of question sheets and answer booklets incurs high paper, ink, and storage costs, contradicting modern environmental sustainability goals.", "Substantial Resource Consumption: ")
    add_bullet("Manual transcription of student scores, paper script bundling, and physical marks entry regularly introduce clerical and calculation errors.", "Vulnerability to Human Error: ")
    add_bullet("Grading hundreds of exam scripts manually takes weeks, postponing grade dissemination and delaying academic progression.", "Significant Delay in Result Publishing: ")
    add_bullet("Client-side timers can be manipulated by pausing JavaScript in developer tools or modifying client device system clocks.", "Vulnerability to Timer Tampering: ")
    add_bullet("If a student's computer loses power or reloads the browser in a basic form, all entered answers are immediately lost, causing extreme student distress.", "Lack of Real-Time State Persistence: ")
    add_bullet("Paper and standard online forms lack built-in monitoring to detect when students switch application tabs or look up external web materials.", "Absence of Active Proctoring: ")
    add_bullet("Existing systems produce simple numerical marks without actionable feedback, preventing students from identifying specific concept deficiencies.", "Generic, Non-Actionable Feedback: ")
    add_bullet("Retrieving past test records requires manual inspection of paper files or scattered spreadsheets, hindering institutional auditing.", "Difficult Record Archiving: ")

    add_heading_2("2.2 PROPOSED SYSTEM")
    add_heading_3("2.2.1 DESCRIPTION")
    add_p("The proposed Online Examination System using Python is designed as a centralized, web-based platform that completely automates the lifecycle of academic evaluations. Developed on a modular Python Flask architecture, the system establishes a secure and authenticated environment where students, teachers, and department heads execute their respective examination roles seamlessly.")
    add_p("The system introduces a server-authoritative timing model where the backend records the official start time and calculates the exact allowable end time for every student attempt. When the server-side remaining duration reaches zero, the attempt status is automatically sealed and graded, completely neutralizing client-side browser timer tampering. In addition, an asynchronous auto-save REST API saves each student response to the database in real time as choices are selected, ensuring complete continuity if a browser window is refreshed or temporarily disconnected.")
    add_p("For administrative users, the platform offers an intuitive question authoring suite supporting multiple question paradigms (single choice, multi-choice with partial scoring, true/false, and short answer) accompanied by instant auto-grading. A specialized Executive Head proctoring dashboard records integrity events (such as window blur, tab switching, and clipboard copy attempts) to ensure academic honesty.")

    add_heading_3("2.2.2 FEATURES")
    add_p("The proposed Online Examination System incorporates a rich set of modern software features:")
    add_bullet("Provides dedicated interfaces and permission boundaries for Student examinees, Instructors/Admins, and Executive Department Heads.", "Multi-Tier Role Architecture: ")
    add_bullet("Prevents exam duration manipulation by computing remaining time on the server backend independently of client system clocks.", "Server-Authoritative Timer: ")
    add_bullet("Background JavaScript triggers asynchronous POST requests to update individual answer selections in SQLite without page reloading.", "Real-Time Asynchronous Auto-Save: ")
    add_bullet("Supports single-choice, multiple-choice (with proportional partial credit), true/false, and short answer formats.", "Multi-Format Question Management: ")
    add_bullet("Categorizes questions by academic topic and difficulty levels (Easy, Medium, Hard) to enable balanced question sets.", "Topic & Difficulty Bank: ")
    add_bullet("Client-side event listeners detect and log browser tab switching, window defocus, and copy attempts, dynamically updating an integrity score.", "Proctoring & Integrity Audit: ")
    add_bullet("Evaluates objective question submissions immediately upon completion, saving instructor time and delivering instant feedback.", "Instant Automated Grading: ")
    add_bullet("Displays comprehensive question-by-question analysis, correct answer justifications, and performance percentages.", "Rich Student Scorecards: ")
    add_bullet("Generates a unique, tamper-resistant digital certificate of achievement with a cryptographic ID for passing students.", "Verifiable Digital Certificates: ")
    add_bullet("Provides executive visualizations showing pass/fail distributions, average score metrics, and suspicious activity logs.", "Executive Analytics & Monitoring: ")
    add_bullet("Protects credentials with PBKDF2-SHA256 password hashing and defends endpoints against CSRF and unauthorized role access.", "Secure Authentication: ")
    add_bullet("Built using modern dark-mode glassmorphism visual elements, clean typography, and responsive mobile-friendly layouts.", "Modern User Interface: ")

    # =========================================================================
    # CHAPTER 3: SYSTEM DESIGN AND DEVELOPMENT
    # =========================================================================
    doc.add_page_break()
    add_heading_1("3. SYSTEM DESIGN AND DEVELOPMENT")
    add_p("System Design and Development is the core technical phase that transforms functional requirements into technical specifications, architectural diagrams, database schemas, and modular code structures. This section presents the comprehensive design of the Online Examination System.")

    add_heading_2("3.1 FILE DESIGN")
    add_p("The application adheres to a modular, production-ready directory layout designed according to the Flask Application Factory and Blueprint pattern. System files are organized logically into distinct packages for models, routes, services, static assets, and templates:")
    
    file_headers = ["File / Directory", "Component Type", "Description"]
    file_rows = [
        ["run.py", "Root Entry Script", "Initializes Flask application and runs the local server instance."],
        ["config.py", "Configuration File", "Stores database paths, secret keys, session parameters, and AI API settings."],
        ["seed.py", "Database Seeder", "Seeds default user accounts (Head, Admin, Student) and initial sample exams."],
        ["app/__init__.py", "Application Factory", "Constructs the Flask app, initializes SQLAlchemy and Flask-Login, registers Blueprints."],
        ["app/models/schema.py", "Data Models", "Defines relational entities: User, Exam, Question, Attempt, Answer, IntegrityEvent."],
        ["app/routes/auth.py", "Blueprint Route", "Handles user registration, authentication login, logout, and session state."],
        ["app/routes/student.py", "Blueprint Route", "Manages student exam listings, exam taking, auto-save API, results, and certificates."],
        ["app/routes/admin.py", "Blueprint Route", "Manages exam creation, question bank authoring, manual publishing, and analytics."],
        ["app/routes/head.py", "Blueprint Route", "Executive dashboard, institution-wide user management, and proctoring audit log."],
        ["app/routes/api.py", "REST API Blueprint", "Handles asynchronous answer auto-saving and real-time integrity event logging."],
        ["app/services/scoring.py", "Business Logic Service", "Executes automated grading algorithms, partial credit logic, and score summaries."],
        ["app/services/ai_service.py", "AI Analytics Service", "Manages automated exam question generation and personalized study feedback."],
        ["app/static/css/style.css", "Design System", "CSS custom properties, glassmorphism card styling, responsive layouts, badges."],
        ["app/static/js/exam.js", "Client Examination Engine", "Handles server timer sync, proctoring event listeners, and async auto-saving."],
        ["app/templates/base.html", "Master Template", "Global HTML structure, navigation bar, flash notifications, and shared styles."],
        ["data/oes.db", "Database File", "SQLite database file storing all institutional examination data."]
    ]
    add_table_data(file_headers, file_rows, col_widths=[1.8, 1.8, 2.9])

    add_heading_2("3.2 INPUT DESIGN")
    add_p("Input design governs how data enters the system, prioritizing user convenience, data consistency, and defensive validation. The primary input interfaces include:")
    
    add_p("1. User Authentication Input:", bold_prefix=None)
    add_bullet("Full Name (Alphabetical string, 2 to 120 characters).", "Name: ")
    add_bullet("Institutional Email Address (Validated RFC-compliant email; acts as unique login key).", "Email: ")
    add_bullet("Account Password (Minimum 6 characters, masked input, hashed via PBKDF2).", "Password: ")
    add_bullet("Role Selector ('student' for examinees, 'admin' for test administrators).", "Role: ")

    add_p("2. Exam Configuration Input:", bold_prefix=None)
    add_bullet("Title of the examination (e.g., 'Python Fundamentals Assessment').", "Exam Title: ")
    add_bullet("Brief syllabus description and instructions to candidates.", "Description: ")
    add_bullet("Allowable examination time in minutes (Integer between 5 and 180).", "Duration: ")
    add_bullet("Minimum percentage required to pass (Float between 0.0 and 100.0).", "Pass Mark: ")
    add_bullet("Publication status toggle ('Published' to make visible to students).", "Published State: ")

    add_p("3. Question Creation Input:", bold_prefix=None)
    add_bullet("Rich text query stating the question premise clearly.", "Question Text: ")
    add_bullet("Dropdown ('single_choice', 'multi_choice', 'true_false', 'short_answer').", "Question Type: ")
    add_bullet("JSON/array of selectable alternatives for choice questions.", "Options List: ")
    add_bullet("Array of correct option values or text key strings.", "Correct Answer(s): ")
    add_bullet("Weightage allocated to the question (Float >= 0.5 points).", "Points: ")
    add_bullet("Subject matter category (e.g., 'Data Structures', 'Functions').", "Topic: ")
    add_bullet("Complexity rating ('easy', 'medium', 'hard').", "Difficulty: ")
    add_bullet("Educational justification explaining why the correct choice is accurate.", "Explanation: ")

    add_p("4. Examination Taking & Auto-Save Input:", bold_prefix=None)
    add_bullet("Current question ID being responded to.", "Question ID: ")
    add_bullet("Selected option radio value, checkbox array, or string response.", "Student Response: ")
    add_bullet("Timestamp of interaction captured by the client script.", "Interaction Time: ")

    add_p("5. Proctoring Event Input:", bold_prefix=None)
    add_bullet("Identifier of suspicious event ('tab_switch', 'window_blur', 'copy_attempt').", "Event Type: ")
    add_bullet("Contextual string detailing the browser event trigger.", "Event Details: ")

    add_heading_2("3.3 OUTPUT DESIGN")
    add_p("Output design describes the information presented to users after processing. Outputs are designed for maximum clarity, actionable insight, and visual excellence:")
    
    add_p("1. Student Dashboard Output:", bold_prefix=None)
    add_bullet("Displays active available assessments with duration, question count, and pass threshold.", "Available Exams Card: ")
    add_bullet("Presents completed attempts, recorded percentage, passed/failed badges, and certificate download links.", "Past Assessments Table: ")

    add_p("2. Live Examination Interface Output:", bold_prefix=None)
    add_bullet("Prominent visual countdown timer updated every second, synchronizing with backend time.", "Real-Time Timer Display: ")
    add_bullet("Real-time badge indicating whether latest responses are securely saved on the server.", "Auto-Save Status Badge: ")
    add_bullet("Interactive grid showing answered, unanswered, and current question statuses.", "Question Navigation Palette: ")

    add_p("3. Scorecard & Performance Summary Output:", bold_prefix=None)
    add_bullet("Final points obtained, maximum available score, and overall percentage.", "Total Score Metric: ")
    add_bullet("Clear visual indicator ('PASSED' in green badge or 'NEEDS IMPROVEMENT' in red badge).", "Pass/Fail Status: ")
    add_bullet("Shows the examinee's final integrity rating (e.g., 95% based on infractions logged).", "Integrity Rating: ")
    add_bullet("Full question-by-question review showing the student's answer, correct answer, awarded points, and detailed explanation.", "Detailed Review Section: ")

    add_p("4. Digital Certificate of Achievement Output:", bold_prefix=None)
    add_bullet("Official certificate displaying student name, examination title, score percentage, date of completion, and unique cryptographic verification ID.", "Academic Certificate: ")

    add_p("5. Administrative & Executive Dashboard Outputs:", bold_prefix=None)
    add_bullet("Aggregate counts of total exams, registered questions, student attempts, and overall pass rates.", "Overview Metric Cards: ")
    add_bullet("Detailed breakdown of student attempts, completion dates, scores, and review options.", "Student Performance Ledger: ")
    add_bullet("Chronological stream of proctoring events detailing candidate infractions, timestamps, and integrity deductions.", "Proctor & Audit Log Table: ")

    add_heading_2("3.4 DATABASE DESIGN")
    add_p("The database is structured in SQLite using third normal form (3NF) principles to ensure data integrity, prevent duplicate data, and maintain clear foreign key relationships.")

    add_p("Table 1: users (Stores user account credentials and roles)")
    user_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    user_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique user identifier."],
        ["name", "VARCHAR(120)", "NOT NULL", "Full legal name of the user."],
        ["email", "VARCHAR(120)", "NOT NULL, UNIQUE, INDEX", "Institutional email for authentication."],
        ["password_hash", "VARCHAR(256)", "NOT NULL", "PBKDF2 salted hash of the password."],
        ["role", "VARCHAR(20)", "NOT NULL, DEFAULT 'student'", "User role: 'student', 'admin', or 'head'."],
        ["created_at", "DATETIME", "DEFAULT UTC NOW", "Account registration timestamp."]
    ]
    add_table_data(user_headers, user_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_p("Table 2: exams (Stores examination metadata and configurations)")
    exam_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    exam_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique exam identifier."],
        ["title", "VARCHAR(200)", "NOT NULL", "Descriptive title of the exam."],
        ["description", "TEXT", "NULLABLE", "Instructions and syllabus summary."],
        ["duration_minutes", "INTEGER", "NOT NULL, DEFAULT 30", "Allowed test time in minutes."],
        ["pass_mark", "FLOAT", "NOT NULL, DEFAULT 50.0", "Minimum passing score percentage."],
        ["is_published", "BOOLEAN", "DEFAULT FALSE", "Visibility toggle for student access."],
        ["created_by_id", "INTEGER", "FOREIGN KEY -> users.id", "User ID of creator (Admin)."],
        ["created_at", "DATETIME", "DEFAULT UTC NOW", "Creation timestamp."]
    ]
    add_table_data(exam_headers, exam_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_p("Table 3: questions (Stores assessment question items)")
    q_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    q_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique question identifier."],
        ["exam_id", "INTEGER", "FOREIGN KEY -> exams.id", "Associated exam ID."],
        ["text", "TEXT", "NOT NULL", "Full text of the question."],
        ["question_type", "VARCHAR(30)", "NOT NULL", "single_choice, multi_choice, true_false, short_answer."],
        ["options_json", "TEXT", "NULLABLE", "JSON-encoded array of selectable options."],
        ["correct_answers_json", "TEXT", "NOT NULL", "JSON-encoded array of correct choices."],
        ["points", "FLOAT", "NOT NULL, DEFAULT 1.0", "Marks weightage allocated."],
        ["explanation", "TEXT", "NULLABLE", "Educational rationale for correct answer."],
        ["topic", "VARCHAR(100)", "NOT NULL, DEFAULT 'General'", "Subject matter topic category."],
        ["difficulty", "VARCHAR(20)", "NOT NULL, DEFAULT 'medium'", "easy, medium, or hard."],
        ["created_at", "DATETIME", "DEFAULT UTC NOW", "Creation timestamp."]
    ]
    add_table_data(q_headers, q_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_p("Table 4: attempts (Stores student exam sessions and final scores)")
    att_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    att_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique attempt session ID."],
        ["user_id", "INTEGER", "FOREIGN KEY -> users.id", "Student candidate ID."],
        ["exam_id", "INTEGER", "FOREIGN KEY -> exams.id", "Exam being taken."],
        ["start_time", "DATETIME", "NOT NULL, DEFAULT UTC NOW", "Backend start timestamp."],
        ["end_time", "DATETIME", "NOT NULL", "Server-authoritative target finish time."],
        ["submitted_at", "DATETIME", "NULLABLE", "Actual student submission time."],
        ["status", "VARCHAR(30)", "NOT NULL, DEFAULT 'in_progress'", "in_progress, submitted, graded."],
        ["score", "FLOAT", "DEFAULT 0.0", "Total marks achieved."],
        ["max_score", "FLOAT", "NOT NULL, DEFAULT 0.0", "Maximum possible exam marks."],
        ["passed", "BOOLEAN", "DEFAULT FALSE", "Whether score >= pass_mark."],
        ["integrity_score", "INTEGER", "DEFAULT 100", "Proctoring score (100% minus penalties)."],
        ["certificate_id", "VARCHAR(64)", "UNIQUE, NULLABLE", "Cryptographic verifiable certificate ID."]
    ]
    add_table_data(att_headers, att_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_p("Table 5: answers (Stores examinee responses for each question)")
    ans_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    ans_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique answer entry ID."],
        ["attempt_id", "INTEGER", "FOREIGN KEY -> attempts.id", "Parent attempt session ID."],
        ["question_id", "INTEGER", "FOREIGN KEY -> questions.id", "Question reference ID."],
        ["student_response_json", "TEXT", "NULLABLE", "JSON-encoded student answer data."],
        ["is_correct", "BOOLEAN", "NULLABLE", "Grading evaluation result."],
        ["points_awarded", "FLOAT", "DEFAULT 0.0", "Calculated marks for this question."],
        ["saved_at", "DATETIME", "DEFAULT UTC NOW", "Last auto-save update timestamp."]
    ]
    add_table_data(ans_headers, ans_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_p("Table 6: integrity_events (Stores proctoring audit log infractions)")
    ev_headers = ["Field Name", "Data Type", "Constraint", "Description"]
    ev_rows = [
        ["id", "INTEGER", "PRIMARY KEY, AUTOINCREMENT", "Unique audit event ID."],
        ["attempt_id", "INTEGER", "FOREIGN KEY -> attempts.id", "Associated exam attempt ID."],
        ["event_type", "VARCHAR(50)", "NOT NULL", "tab_switch, window_blur, copy_attempt."],
        ["details", "TEXT", "NULLABLE", "Contextual incident details."],
        ["timestamp", "DATETIME", "DEFAULT UTC NOW", "Exact incident timestamp."]
    ]
    add_table_data(ev_headers, ev_rows, col_widths=[1.5, 1.5, 2.0, 1.5])

    add_heading_2("3.5 SYSTEM DEVELOPMENT")
    add_p("System development is the practical engineering process of translating the system design specifications into maintainable, well-structured software code. The Online Examination System was developed using Python Flask following an iterative Model-View-Controller (MVC) architectural design:")
    add_bullet("Encapsulated via Flask-SQLAlchemy models in 'app/models/schema.py', handling data definitions, relationships, constraints, and persistence methods.", "The Model Layer: ")
    add_bullet("Implemented via Jinja2 HTML5 templates in 'app/templates/', styled with custom CSS3 glassmorphic design variables, providing clean and responsive user screens.", "The View Layer: ")
    add_bullet("Structured across discrete Flask Blueprints in 'app/routes/', handling HTTP requests, session validation, form parsing, service orchestration, and response redirection.", "The Controller Layer: ")
    add_p("A dedicated Service Layer ('app/services/') decouples business logic from HTTP route handlers. For example, 'scoring.py' encapsulates grading mathematics, and 'ai_service.py' handles AI integration. This clean separation of concerns facilitates automated testing and seamless future extensions.")

    add_heading_2("3.6 DESCRIPTION OF MODULES")
    add_p("The Online Examination System comprises eight core functional modules that work harmoniously:")

    add_p("1. Authentication Module (app/routes/auth.py):", bold_prefix=None)
    add_p("Manages user identity and access control. It provides registration for students and admins, secure login with PBKDF2 password verification, role identification, and secure session teardown on logout. Unauthenticated users attempting to access protected endpoints are automatically redirected with informative flash messages.")

    add_p("2. Student Examination Module (app/routes/student.py):", bold_prefix=None)
    add_p("Provides students with their personalized assessment dashboard, displays available exams, initiates new exam attempts, renders the live testing interface, and displays finalized scorecards and downloadable certificates. It interfaces directly with the server-authoritative timer to enforce examination boundaries.")

    add_p("3. Administrative Exam & Question Module (app/routes/admin.py):", bold_prefix=None)
    add_p("Empowers educators and administrators to create, modify, and publish examinations. It features an interactive question authoring suite where teachers can compose single-choice, multiple-choice, true/false, or short-answer questions, allocate point weightages, specify explanations, and categorize items by topic and difficulty.")

    add_p("4. Executive Head & Audit Module (app/routes/head.py):", bold_prefix=None)
    add_p("Provides executive oversight for Department Heads and Chief Examiners. It includes institution-wide user management (viewing, promoting, or modifying user roles) and a live Proctoring Audit Log that displays real-time security events captured across all examination attempts.")

    add_p("5. Asynchronous API & Proctoring Module (app/routes/api.py):", bold_prefix=None)
    add_p("Supplies lightweight JSON REST endpoints. The auto-save endpoint accepts asynchronous POST payloads from the client exam script to record answers in the database without page refreshes. The integrity endpoint logs client proctoring infractions and updates the candidate's session integrity score.")

    add_p("6. Automated Scoring Service (app/services/scoring.py):", bold_prefix=None)
    add_p("Executes server-side grading algorithms upon examination submission. For single-choice and true/false questions, it checks exact matches. For multiple-choice questions, it calculates proportional partial credit for partially correct option combinations while penalizing erroneous selections. For short answers, it performs case-insensitive normalized string matching.")

    add_p("7. Vector Search Service (app/services/vector_service.py):", bold_prefix=None)
    add_p("Integrates an embedded ChromaDB vector database to store mathematical vector embeddings of questions. It enables administrators to search question banks semantically and prevents duplicate question entry across assessments.")

    add_p("8. AI Analysis Service (app/services/ai_service.py):", bold_prefix=None)
    add_p("Leverages the OpenRouter API (Google Gemini 2.5 Flash) to generate automated assessment questions from teacher prompts and generate personalized feedback summaries highlighting individual student strengths and study recommendations.")

    # =========================================================================
    # CHAPTER 4: TESTING AND IMPLEMENTATION
    # =========================================================================
    doc.add_page_break()
    add_heading_1("4. TESTING AND IMPLEMENTATION")
    add_p("Testing and Implementation represents the verification and rollout stage of software development. It verifies that all functional components perform according to specification, handle erroneous inputs gracefully, safeguard system integrity, and install reliably in production environments.")

    add_heading_2("4.1 TEST CASES EXAMPLE")
    add_p("The system was thoroughly validated using both automated Pytest suites and manual end-to-end verification. The following structured test matrix details key test cases executed across all functional modules:")

    tc_headers = ["Test Case", "Module", "Test Scenario & Input", "Expected Result", "Status"]
    tc_rows = [
        ["TC01", "Auth", "Submit valid email & password", "Authenticate user and open dashboard", "Pass"],
        ["TC02", "Auth", "Submit incorrect password", "Deny access and display error alert", "Pass"],
        ["TC03", "Auth", "Register existing email", "Reject registration with unique constraint warning", "Pass"],
        ["TC04", "Admin", "Create new examination with title & duration", "Exam saved to database in draft state", "Pass"],
        ["TC05", "Admin", "Add single-choice question with 4 options", "Question stored with correct answer index", "Pass"],
        ["TC06", "Admin", "Add multi-choice question with multiple correct options", "Options and answers serialized to JSON correctly", "Pass"],
        ["TC07", "Student", "Start published examination", "Create attempt row; compute server end_time", "Pass"],
        ["TC08", "Student", "Select answer option in exam", "Asynchronous API auto-saves response with 200 OK", "Pass"],
        ["TC09", "Student", "Browser page refreshed during live exam", "Reload loads previously selected responses intact", "Pass"],
        ["TC10", "Security", "Client device system clock adjusted forward", "Server timer ignores client clock and preserves true time", "Pass"],
        ["TC11", "Proctor", "Examinee switches browser tab", "Capture window blur, log integrity infraction event", "Pass"],
        ["TC12", "Scoring", "Submit exam before timer expiry", "Calculate score, mark attempt 'submitted', grade answers", "Pass"],
        ["TC13", "Scoring", "Server timer reaches zero without student submission", "Server auto-submits and finalizes attempt", "Pass"],
        ["TC14", "Student", "Score >= pass_mark threshold", "Mark attempt passed and issue unique certificate ID", "Pass"],
        ["TC15", "Head", "Executive views proctoring audit log", "Display chronological events with integrity scores", "Pass"]
    ]
    add_table_data(tc_headers, tc_rows, col_widths=[0.8, 0.9, 2.5, 2.0, 0.6])

    add_heading_2("4.2 IMPLEMENTATION")
    add_p("Implementation involves converting the software design into a functioning, operational application deployed within a target execution environment. The implementation of the Online Examination System followed a systematic five-stage deployment methodology:")
    
    add_p("Stage 1: Python Environment & Dependency Isolation:", bold_prefix=None)
    add_p("A dedicated Python virtual environment is established to ensure dependency isolation. Core dependencies specified in 'requirements.txt' (Flask, Flask-SQLAlchemy, Flask-Login, Werkzeug, ChromaDB, Requests) are installed using the pip package manager.")

    add_p("Stage 2: Database Schema Migration & Seeding:", bold_prefix=None)
    add_p("The SQLite relational database schema is created by invoking 'db.create_all()' within the Flask application context. The database seeding script ('seed.py') is executed to populate default accounts (Head Examiner: 'head@oes.com', Administrator: 'admin@oes.com', Student: 'student@oes.com') and initial practice examinations.")

    add_p("Stage 3: Blueprint Routing & Template Binding:", bold_prefix=None)
    add_p("Flask Blueprints ('auth_bp', 'student_bp', 'admin_bp', 'head_bp', 'api_bp') are registered on the central application instance. Jinja2 templates are configured with inheritance hierarchies extending 'base.html', ensuring uniform navigation bars, badge styling, and mobile-friendly responsive wrappers across all views.")

    add_p("Stage 4: Client-Side Engine & Proctor Event Hooks:", bold_prefix=None)
    add_p("The client-side JavaScript engine ('app/static/js/exam.js') is bound to the examination template. Event listeners for 'visibilitychange', 'blur', and 'copy' are activated, linking browser events with the backend API endpoint ('/api/integrity-event'). The auto-save debouncer is initialized to asynchronously sync radio and checkbox state changes.")

    add_p("Stage 5: Security Hardening & Session Configuration:", bold_prefix=None)
    add_p("Session cookies are configured with HttpOnly and SameSite attributes to protect against session hijacking and cross-site scripting (XSS). Server-side route decorators ('@login_required', '@admin_required', '@head_required') are enforced on all administrative endpoints.")

    add_heading_2("4.3 DEPLOYMENT OPTIONS")
    add_p("The Online Examination System is architected to accommodate diverse institutional hosting requirements through three primary deployment models:")

    add_p("1. Campus Local Area Network (LAN) / On-Premise Deployment:", bold_prefix=None)
    add_bullet("Best suited for single-campus testing centers and laboratory environments.", "Target Environment: ")
    add_bullet("The application is hosted on a central departmental server running Windows Server or Ubuntu Linux.", "Architecture: ")
    add_bullet("Employs the production-grade Waitress WSGI server on Windows ('waitress-serve --port=5000 run:app') or Gunicorn on Linux.", "Web Server: ")
    add_bullet("Uses local SQLite storage with regular automated file backups; requires zero external internet connectivity for core exam taking.", "Data Storage: ")
    add_bullet("Complete institutional data privacy, minimal latency, and full functionality even during external internet blackouts.", "Key Advantage: ")

    add_p("2. Containerized Cloud Deployment (Docker & Nginx):", bold_prefix=None)
    add_bullet("Suited for multi-campus colleges and universities conducting large-scale simultaneous testing.", "Target Environment: ")
    add_bullet("The application, Python runtime, and dependencies are packaged into an immutable Docker container image.", "Architecture: ")
    add_bullet("Nginx acts as a high-performance reverse proxy handling SSL/TLS certificate termination and static asset caching, routing dynamic requests to multiple Gunicorn WSGI worker processes.", "Web Server: ")
    add_bullet("SQLite can be upgraded to an external PostgreSQL database container by simply updating the SQLALCHEMY_DATABASE_URI in 'config.py'.", "Data Storage: ")
    add_bullet("High horizontal scalability, fault tolerance, and predictable containerized behavior across any cloud platform (AWS, GCP, Azure).", "Key Advantage: ")

    add_p("3. Managed Platform-as-a-Service (PaaS) Cloud Deployment:", bold_prefix=None)
    add_bullet("Ideal for academic institutions seeking zero server hardware management overhead.", "Target Environment: ")
    add_bullet("Direct deployment on cloud application platforms such as Render, Railway, Heroku, or PythonAnywhere directly from Git repositories.", "Architecture: ")
    add_bullet("Automatic SSL/TLS provisioning, continuous deployment upon Git push, and managed database backups.", "Key Advantage: ")

    # =========================================================================
    # CHAPTER 5: CONCLUSION
    # =========================================================================
    doc.add_page_break()
    add_heading_1("CONCLUSION")
    add_p("The Online Examination System using Python successfully addresses the longstanding limitations of traditional paper-based examinations and rudimentary online survey forms. By integrating a modular Python Flask backend with a responsive, glassmorphism-styled user interface and an SQLite transactional database, the project delivers a robust, secure, and modern academic assessment platform.")
    add_p("Through the implementation of a server-authoritative timer, the system eliminates client-side duration tampering, ensuring fair and standardized test conditions for all examinees. The asynchronous background auto-save mechanism provides resilience against network fluctuations, ensuring that student answers are permanently preserved. Furthermore, the platform's multi-format question support, instant automated grading, verifiable digital certificates, and comprehensive proctoring audit log provide tremendous value to students, faculty members, and institutional administrators.")
    add_p("The modular architecture ensures that future technological enhancements can be integrated effortlessly. Planned enhancements include:")
    add_bullet("Integrating computer vision models (OpenCV / MediaPipe) to detect multiple faces, absence from the screen, or unauthorized smartphone usage via the student's webcam.", "Webcam AI Proctoring: ")
    add_bullet("Enabling math notation rendering (KaTeX / MathJax) and programming code execution sandboxes for computer science practical exams.", "Rich Media & Code Assessment: ")
    add_bullet("Developing a lightweight mobile client using Flutter or Progressive Web App (PWA) standards for offline-first testing.", "Mobile Application: ")
    add_bullet("Automating SMS and email notifications to alert students to newly published examinations and grade publications.", "Notification Gateway: ")
    add_p("In conclusion, the developed system demonstrates how modern software engineering principles and Python web technologies can transform institutional examination management into an efficient, trustworthy, and paperless operational model.")

    # =========================================================================
    # CHAPTER 6: BIBLIOGRAPHY
    # =========================================================================
    doc.add_page_break()
    add_heading_1("BIBLIOGRAPHY")
    add_p("The following textbooks, official documentation, academic papers, and technical standards were consulted during the design, development, and documentation of the Online Examination System:")
    
    add_heading_2("6.1 BOOKS")
    add_bullet("Pressman, Roger S., and Bruce R. Maxim. Software Engineering: A Practitioner's Approach. 9th ed., McGraw-Hill Education, 2020.", "1. ")
    add_bullet("Sommerville, Ian. Software Engineering. 10th ed., Pearson Education, 2016.", "2. ")
    add_bullet("Grinberg, Miguel. Flask Web Development: Developing Web Applications with Python. 2nd ed., O'Reilly Media, 2018.", "3. ")
    add_bullet("Silberschatz, Abraham, Henry F. Korth, and S. Sudarshan. Database System Concepts. 7th ed., McGraw-Hill Education, 2019.", "4. ")
    add_bullet("Lutz, Mark. Learning Python. 5th ed., O'Reilly Media, 2013.", "5. ")

    add_heading_2("6.2 ONLINE RESOURCES & OFFICIAL DOCUMENTATION")
    add_bullet("Python Software Foundation. Python 3.12 Documentation. Available at: https://docs.python.org/3/", "1. ")
    add_bullet("Pallets Projects. Flask Documentation (Version 3.0.x). Available at: https://flask.palletsprojects.com/", "2. ")
    add_bullet("SQLAlchemy Authors. SQLAlchemy 2.0 Documentation. Available at: https://docs.sqlalchemy.org/", "3. ")
    add_bullet("Mozilla Developer Network (MDN). MDN Web Docs: HTML5, CSS3, and JavaScript APIs. Available at: https://developer.mozilla.org/", "4. ")
    add_bullet("Chroma Core Team. ChromaDB Vector Database Documentation. Available at: https://docs.trychroma.com/", "5. ")
    add_bullet("Pytest Development Team. Pytest: Robust Python Testing. Available at: https://docs.pytest.org/", "6. ")

    add_heading_2("6.3 DEVELOPMENT TOOLS & TECHNICAL STANDARDS")
    add_bullet("Visual Studio Code (VS Code) – Integrated source code editor with Python extensions.", "1. ")
    add_bullet("Google Chrome & Microsoft Edge Developer Tools – Used for client performance analysis and DOM verification.", "2. ")
    add_bullet("OWASP Foundation. OWASP Top 10 Web Application Security Risks. Available at: https://owasp.org/www-project-top-ten/", "3. ")

    # =========================================================================
    # CHAPTER 7: APPENDICES
    # =========================================================================
    doc.add_page_break()
    add_heading_1("APPENDICES")

    add_heading_2("A) SAMPLE CODING")
    add_p("The following production code excerpts demonstrate key architectural components of the Online Examination System:")

    add_p("Excerpt 1: Server-Authoritative Timer & Student Exam Session (app/routes/student.py)", bold_prefix=None)
    code_1 = """@student_bp.route('/exam/<int:exam_id>/take')
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
    
    # Server-Authoritative Time Calculation
    now = datetime.now(timezone.utc)
    remaining_seconds = max(0, int((attempt.end_time - now).total_seconds()))
    if remaining_seconds <= 0:
        return redirect(url_for('student.submit_exam', attempt_id=attempt.id))
        
    return render_template('student/take_exam.html', exam=exam, attempt=attempt, remaining_seconds=remaining_seconds)"""
    add_code_block(code_1)

    add_p("Excerpt 2: Automated Grading & Partial Credit Scoring Algorithm (app/services/scoring.py)", bold_prefix=None)
    code_2 = """def score_attempt(attempt):
    total_score = 0.0
    for answer in attempt.answers:
        q = answer.question
        if q.question_type == 'single_choice' or q.question_type == 'true_false':
            is_correct = (str(answer.student_response).strip() == str(q.correct_answers[0]).strip())
            answer.is_correct = is_correct
            answer.points_awarded = q.points if is_correct else 0.0
        elif q.question_type == 'multi_choice':
            student_set = set(answer.student_response or [])
            correct_set = set(q.correct_answers)
            if student_set == correct_set:
                answer.is_correct = True
                answer.points_awarded = q.points
            elif student_set.issubset(correct_set) and len(student_set) > 0:
                answer.is_correct = False
                answer.points_awarded = round(q.points * (len(student_set) / len(correct_set)), 2)
            else:
                answer.is_correct = False
                answer.points_awarded = 0.0
        total_score += answer.points_awarded
    attempt.score = total_score
    attempt.passed = (attempt.score / attempt.max_score * 100 >= attempt.exam.pass_mark) if attempt.max_score > 0 else False
    attempt.status = 'graded'
    db.session.commit()
    return attempt"""
    add_code_block(code_2)

    add_p("Excerpt 3: Client-Side Auto-Save & Proctor Event Listener (app/static/js/exam.js)", bold_prefix=None)
    code_3 = """// Asynchronous Answer Auto-Save
function saveAnswer(questionId, responseValue) {
    const statusBadge = document.getElementById('save-status');
    statusBadge.innerText = 'Saving...';
    
    fetch('/api/save-answer', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            attempt_id: ATTEMPT_ID,
            question_id: questionId,
            response: responseValue
        })
    }).then(res => res.json()).then(data => {
        if (data.success) {
            statusBadge.innerText = 'All responses saved';
        }
    }).catch(err => {
        statusBadge.innerText = 'Offline - will retry';
    });
}

// Proctoring Window Blur & Tab Switch Listener
window.addEventListener('visibilitychange', () => {
    if (document.hidden) {
        fetch('/api/integrity-event', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                attempt_id: ATTEMPT_ID,
                event_type: 'tab_switch',
                details: 'Candidate switched away from examination window'
            })
        });
    }
});"""
    add_code_block(code_3)

    add_heading_2("B) SAMPLE INPUT")
    add_p("This section illustrates the structure and validation rules of sample inputs handled by the system:")
    
    add_p("1. Sample User Registration Input:", bold_prefix=None)
    sample_input_1 = """[POST] /auth/register
Payload:
{
    "name": "M. Logeshwari",
    "email": "logeshwari@college.edu",
    "password": "SecurePassword#2026",
    "role": "student"
}
Validation: Checks email uniqueness, password hash generation, and assigns standard student role."""
    add_code_block(sample_input_1)

    add_p("2. Sample Exam Creation Input:", bold_prefix=None)
    sample_input_2 = """[POST] /admin/exam/new
Payload:
{
    "title": "Python Programming and Data Structures Evaluation",
    "description": "Mid-term comprehensive exam covering data types, control flow, functions, and OOP.",
    "duration_minutes": 45,
    "pass_mark": 60.0,
    "is_published": true
}
Validation: Validates title length, ensures duration is a positive integer, sets published flag."""
    add_code_block(sample_input_2)

    add_p("3. Sample Multi-Choice Question Input:", bold_prefix=None)
    sample_input_3 = """[POST] /admin/exam/1/question/new
Payload:
{
    "text": "Which of the following data types in Python are mutable? (Select all that apply)",
    "question_type": "multi_choice",
    "options": ["List", "Tuple", "Dictionary", "String"],
    "correct_answers": ["List", "Dictionary"],
    "points": 2.0,
    "topic": "Python Data Structures",
    "difficulty": "medium",
    "explanation": "Lists and dictionaries can be modified in place; tuples and strings are immutable."
}
Validation: Serializes options and answers into JSON arrays; confirms non-empty choices."""
    add_code_block(sample_input_3)

    add_p("4. Sample Asynchronous Auto-Save Input:", bold_prefix=None)
    sample_input_4 = """[POST] /api/save-answer
Payload:
{
    "attempt_id": 104,
    "question_id": 12,
    "response": ["List", "Dictionary"]
}
Validation: Confirms attempt is active; updates or inserts answer record in SQLite transaction."""
    add_code_block(sample_input_4)

    add_heading_2("C) SAMPLE OUTPUT")
    add_p("This section illustrates typical outputs generated across different interfaces of the application:")

    add_p("1. Sample Examination Scorecard Output (Web View):", bold_prefix=None)
    sample_out_1 = """================================================================================
                    ONLINE EXAMINATION SCORECARD
================================================================================
Candidate Name     : M. Logeshwari (Degree: 3-B.Sc Computer Science)
Examination Title  : Python Programming and Data Structures Evaluation
Date Completed     : 2026-09-19 10:45:00 UTC
Session Duration   : 42 minutes 18 seconds
--------------------------------------------------------------------------------
FINAL SCORE        : 26.0 / 30.0 pts  (86.7%)
PASS THRESHOLD     : 60.0%
ASSESSMENT RESULT  : PASSED [STATUS: ACCREDITED]
INTEGRITY RATING   : 100% [Zero suspicious infractions recorded]
CERTIFICATE ID     : CERT-OES-2026-88A4DF
--------------------------------------------------------------------------------
TOPIC MASTERY BREAKDOWN:
  * Control Structures     : 100% Mastery (High)
  * Object-Oriented Design : 85% Mastery (High)
  * Exception Handling     : 75% Mastery (Medium)
================================================================================
VERIFIABLE DIGITAL CERTIFICATE GENERATED AND READY FOR EXPORT."""
    add_code_block(sample_out_1)

    add_p("2. Sample Digital Certificate of Achievement (Text Representation):", bold_prefix=None)
    sample_out_2 = """+------------------------------------------------------------------------------+
|             PACHAMUTHU COLLEGE OF ARTS AND SCIENCE FOR WOMEN                 |
|                       DEPARTMENT OF COMPUTER SCIENCE                         |
|                                                                              |
|                    CERTIFICATE OF ACADEMIC ACHIEVEMENT                       |
|                                                                              |
|  This is to certify that                                                     |
|                               M. LOGESHWARI                                  |
|                                                                              |
|  has successfully taken and passed the institutional examination:            |
|             ONLINE EXAMINATION SYSTEM - PYTHON PROGRAMMING                   |
|                                                                              |
|  Achieving a distinction score of 86.7% with an Integrity Rating of 100%.    |
|                                                                              |
|  Issue Date: September 19, 2026                 Certificate ID: CERT-OES-88  |
|  Academic Guide: Ms. C. MOHANAPRIYA, M.Sc.      Principal: Dr. J. THAVAMANI  |
+------------------------------------------------------------------------------+"""
    add_code_block(sample_out_2)

    add_p("3. Sample Executive Proctoring Audit Log Output:", bold_prefix=None)
    sample_out_3 = """================================================================================
                   EXECUTIVE PROCTORING & INTEGRITY AUDIT LOG
================================================================================
Log Timestamp       Attempt ID   Candidate Email      Infraction Event   Integrity
2026-09-19 10:12:04  ATT-102      student2@oes.com     tab_switch         90%
2026-09-19 10:14:22  ATT-102      student2@oes.com     copy_attempt       80%
2026-09-19 10:30:11  ATT-104      logeshwari@oes.com   none (clean)       100%
================================================================================
Summary: 1 active flagged session; 0 critical anomalies detected."""
    add_code_block(sample_out_3)

    # Save document
    filename = "Online_Examination_System_Project_Report.docx"
    doc.save(filename)
    print(f"Document successfully created and saved to '{filename}'.")

if __name__ == '__main__':
    build_complete_project_document()
