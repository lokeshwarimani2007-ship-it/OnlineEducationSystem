import docx

doc = docx.Document('Online_Examination_System_76_Pages_Report.docx')
xml_content = doc._body._element.xml
breaks = xml_content.count('w:type="page"')
print(f"Total Hard Page Breaks: {breaks}")
print(f"Exact Document Page Count: {breaks + 1} Pages")
print(f"Total Paragraphs: {len(doc.paragraphs)}")
print(f"Total Tables: {len(doc.tables)}")
