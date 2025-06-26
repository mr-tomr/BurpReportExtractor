# BurpReportExtractor

# Burp Suite HTML to DOCX Converter

This Python script parses a Burp Suite HTML report and extracts key vulnerability data into a clean, professional DOCX file.

It includes:
- ✅ Issue titles
- ✅ Affected URLs (inferred from HTTP requests)
- ✅ Raw HTTP request/response pairs
- ✅ Red-highlighted text matching Burp's HTML highlights
- ✅ Clean formatting for use in reports or presentations

---

## 📦 Features

- Converts `.html` export from Burp Suite to `.docx`
- Preserves red highlight (`<span class="HIGHLIGHT">`) from HTML
- Uses monospace font and boxed layout for requests/responses
- Supports easy formatting and professional output

---

## 🛠 Usage

```bash
python burp_to_docx.py <input_burp_report.html> <output_report.docx>
