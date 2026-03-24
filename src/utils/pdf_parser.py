import os
from docling.document_converter import DocumentConverter

def convert_pdfs_to_markdown(raw_dir, md_dir):
    converter = DocumentConverter()
    
    if not os.path.exists(md_dir):
        os.makedirs(md_dir)

    for filename in os.listdir(raw_dir):
        if filename.endswith(".pdf"):
            print(f"Parsing: {filename}...")
            pdf_path = os.path.join(raw_dir, filename)
            
            # Convert PDF to Markdown
            result = converter.convert(pdf_path)
            md_content = result.document.export_to_markdown()
            
            # Save Markdown file
            output_path = os.path.join(md_dir, filename.replace(".pdf", ".md"))
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(md_content)

if __name__ == "__main__":
    convert_pdfs_to_markdown("data/raw", "data/markdown")