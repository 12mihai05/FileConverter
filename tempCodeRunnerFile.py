def convert_docx_to_pdf(docx_file_path, pdf_file_path):
    doc = Document(docx_file_path)
    c = canvas.Canvas(pdf_file_path, pagesize=letter)
    y = 750
    for paragraph in doc.paragraphs:
        cleaned_text = clean_text(paragraph.text)
        c.drawString(100, y, cleaned_text)
        y -= 15
    c.save()
    print(f"File {docx_file_path} converted to PDF and saved to {pdf_file_path}.")