
# Exports Requests and Responses from Burp.html reports.
# Exports in to Word format.


import sys
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor

def extract_url_from_request(text):
    lines = text.strip().splitlines()
    if lines:
        parts = lines[0].split()
        if len(parts) > 1:
            return parts[1]
    return "N/A"

def process_html_to_docx(input_html, output_docx):
    with open(input_html, "r", encoding="utf-8") as file:
        soup = BeautifulSoup(file, "html.parser")

    doc = Document()
    doc.add_heading('Burp Suite Report - Requests and Responses', level=1)

    titles = soup.find_all("div", class_="BODH0")
    rr_blocks = soup.find_all("div", class_="rr_div")

    for i in range(0, len(rr_blocks) - 1, 2):
        title_text = titles[i // 2].get_text(strip=True) if i // 2 < len(titles) else f"Issue {i // 2 + 1}"
        request_div = rr_blocks[i]
        response_div = rr_blocks[i + 1]

        request_text = request_div.get_text().strip()
        affected_url = extract_url_from_request(request_text)

        doc.add_paragraph(f"Issue: {title_text}", style='Normal')
        doc.add_paragraph(f"Affected URL: {affected_url}", style='Normal')

        # Request heading
        req_heading = doc.add_paragraph()
        req_run = req_heading.add_run("Request")
        req_run.bold = True
        req_run.font.size = Pt(14)
        req_run.font.color.rgb = RGBColor(0, 0, 0)

        table_r = doc.add_table(rows=1, cols=1)
        table_r.style = 'Table Grid'
        cell_r = table_r.cell(0, 0)
        para_r = cell_r.paragraphs[0]
        for element in request_div.descendants:
            if isinstance(element, str):
                para_r.add_run(element).font.name = 'Courier New'
            elif element.name == "span" and "HIGHLIGHT" in element.get("class", []):
                run = para_r.add_run(element.get_text())
                run.font.name = 'Courier New'
                run.font.color.rgb = RGBColor(255, 0, 0)

        # Response heading
        res_heading = doc.add_paragraph()
        res_run = res_heading.add_run("Response")
        res_run.bold = True
        res_run.font.size = Pt(14)
        res_run.font.color.rgb = RGBColor(0, 0, 0)

        table_s = doc.add_table(rows=1, cols=1)
        table_s.style = 'Table Grid'
        cell_s = table_s.cell(0, 0)
        para_s = cell_s.paragraphs[0]
        for element in response_div.descendants:
            if isinstance(element, str):
                para_s.add_run(element).font.name = 'Courier New'
            elif element.name == "span" and "HIGHLIGHT" in element.get("class", []):
                run = para_s.add_run(element.get_text())
                run.font.name = 'Courier New'
                run.font.color.rgb = RGBColor(255, 0, 0)

    doc.save(output_docx)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python burp_to_docx.py <input_html_file> <output_docx_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    process_html_to_docx(input_file, output_file)
