import os
import csv
import mysql.connector
import pptx
import PyPDF2
from datetime import datetime
import re
import fitz  # PyMuPDF

# ========================= DATABASE CONNECTION =========================
class MySQLStorage:
    def __init__(self, extractors, host="localhost", user="harshit", password="harshit", database="extracted_database3"):
        self.extractors = extractors if isinstance(extractors, list) else [extractors]
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.conn = self._connect_db()
        self._initialize_db()

    def _connect_db(self):
        return mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database
        )

    def _initialize_db(self):
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ExtractedText (
                id INT AUTO_INCREMENT PRIMARY KEY,
                source VARCHAR(50),
                content LONGTEXT,
                UNIQUE KEY (source, content(255))
            )
            """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Tables (
                id INT AUTO_INCREMENT PRIMARY KEY,
                source VARCHAR(50),
                table_data LONGTEXT,
                UNIQUE KEY (source, table_data(255))
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Hyperlinks (
                id INT AUTO_INCREMENT PRIMARY KEY,
                source VARCHAR(50),
                page INT,
                text TEXT,
                url TEXT,
                UNIQUE KEY (source, page, url(255))
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Images (
                id INT AUTO_INCREMENT PRIMARY KEY,
                source VARCHAR(50),
                page INT,
                path TEXT,
                format VARCHAR(10),
                resolution VARCHAR(20),
                UNIQUE KEY (source, path(255))
            )
        """)
        self.conn.commit()
        cursor.close()

    def store_data(self):
        cursor = self.conn.cursor()
        for extractor in self.extractors:
            source_type = extractor.get_source_type()
            
            # Store text
            extracted_text = extractor.extract_text()
            cursor.execute("INSERT IGNORE INTO ExtractedText (source, content) VALUES (%s, %s)", 
                          (source_type, extracted_text))
            
            # Store tables
            for table in extractor.extract_tables():
                table_string = "\n".join(["\t".join(row) for row in table["data"]])
                cursor.execute("INSERT IGNORE INTO Tables (source, table_data) VALUES (%s, %s)", 
                              (source_type, table_string))
            
            # Store hyperlinks
            for page, text, url in extractor.extract_hyperlinks():
                cursor.execute("INSERT IGNORE INTO Hyperlinks (source, page, text, url) VALUES (%s, %s, %s, %s)", 
                              (source_type, page, text, url))
            
            # Store image metadata
            for page, path, fmt, resolution in extractor.extract_images():
                cursor.execute("INSERT IGNORE INTO Images (source, page, path, format, resolution) VALUES (%s, %s, %s, %s, %s)", 
                              (source_type, page, path, fmt, resolution))
        
        self.conn.commit()
        cursor.close()
        self.conn.close()

# ========================= FILE LOADERS =========================
class PPTLoader:
    def __init__(self, file_path):
        self.file_path = file_path
        if not file_path.endswith('.pptx'):
            raise ValueError("Invalid PPT file")
    
    def load_file(self):
        return pptx.Presentation(self.file_path)

class PDFLoader:
    def __init__(self, file_path):
        self.file_path = file_path
        if not file_path.endswith('.pdf'):
            raise ValueError("Invalid PDF file")
    
    def load_file_pypdf(self):
        return PyPDF2.PdfReader(open(self.file_path, 'rb'))
    
    def load_file_pymupdf(self):
        return fitz.open(self.file_path)

# ========================= DATA EXTRACTION =========================
class DataExtractor:
    def get_source_type(self):
        raise NotImplementedError("Subclasses must implement get_source_type method")
    
    def extract_text(self):
        raise NotImplementedError("Subclasses must implement extract_text method")
    
    def extract_hyperlinks(self):
        raise NotImplementedError("Subclasses must implement extract_hyperlinks method")
    
    def extract_images(self, output_dir):
        raise NotImplementedError("Subclasses must implement extract_images method")
    
    def extract_tables(self):
        raise NotImplementedError("Subclasses must implement extract_tables method")

class PPTDataExtractor(DataExtractor):
    def __init__(self, loader):
        self.loader = loader.load_file()
        self.file_path = loader.file_path

    def get_source_type(self):
        return "PPT"

    def extract_text(self):
        extracted_text = []
        for i, slide in enumerate(self.loader.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    extracted_text.append(f"Page {i}: {shape.text}")
        return "\n".join(extracted_text)

    def extract_hyperlinks(self):
        hyperlinks = []
        for i, slide in enumerate(self.loader.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.hyperlink and run.hyperlink.address:
                                hyperlinks.append((i, run.text, run.hyperlink.address))
        return hyperlinks

    def extract_images(self, output_dir="output/ppt_images"):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        images = []
        for i, slide in enumerate(self.loader.slides, 1):
            for j, shape in enumerate(slide.shapes):
                if shape.shape_type == 13:  # Shape type 13 corresponds to images
                    image = shape.image
                    image_format = image.ext
                    width_px = int(shape.width / 9525)
                    height_px = int(shape.height / 9525)
                    resolution = f"{width_px}x{height_px}"
                    image_filename = f"{output_dir}/slide_{i}_img_{j+1}.{image_format}"
                    with open(image_filename, "wb") as f:
                        f.write(image.blob)
                    images.append((i, image_filename, image_format, resolution))
        return images

    def extract_tables(self):
        tables = []
        for i, slide in enumerate(self.loader.slides, 1):
            for shape in slide.shapes:
                if hasattr(shape, "table"):
                    table_data = []
                    row_count = len(shape.table.rows)
                    col_count = len(shape.table.columns)
                    for row in shape.table.rows:
                        table_data.append([cell.text for cell in row.cells])
                    tables.append({"page": i, "data": table_data, "size": f"{row_count}x{col_count}"})
        return tables

class PDFDataExtractor(DataExtractor):
    def __init__(self, loader):
        self.file_path = loader.file_path
        self.pdf_reader = loader.load_file_pypdf()
        self.pdf_doc = loader.load_file_pymupdf()

    def get_source_type(self):
        return "PDF"

    def extract_text(self):
        extracted_text = []
        for i, page in enumerate(self.pdf_reader.pages, 1):
            text = page.extract_text()
            if text:
                extracted_text.append(f"Page {i}: {text}")
        return "\n".join(extracted_text)
    
    def extract_hyperlinks(self):
        hyperlinks = []
        for i, page in enumerate(self.pdf_doc, 1):
            links = page.get_links()
            for link in links:
                if 'uri' in link:
                    # Get text near the link
                    rect = fitz.Rect(link['from'])
                    # Extend the rect by a small amount to capture text
                    rect.x0 -= 5
                    rect.y0 -= 5
                    rect.x1 += 5
                    rect.y1 += 5
                    
                    # Get words within the extended rectangle
                    words = page.get_text("words", clip=rect)
                    link_text = " ".join([word[4] for word in words]) if words else "Unknown link text"
                    hyperlinks.append((i, link_text, link['uri']))
        return hyperlinks

    def extract_images(self, output_dir="output/pdf_images"):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        images = []
        for i, page in enumerate(self.pdf_doc, 1):
            image_list = page.get_images(full=True)
            for j, img in enumerate(image_list, 1):
                xref = img[0]
                base_image = self.pdf_doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                image_filename = f"{output_dir}/page_{i}_img_{j}.{image_ext}"
                with open(image_filename, "wb") as f:
                    f.write(image_bytes)
                
                # Get dimensions from the image if available
                width = base_image.get("width", 0)
                height = base_image.get("height", 0)
                resolution = f"{width}x{height}" if width and height else "unknown"
                images.append((i, image_filename, image_ext, resolution))
        return images

    def extract_tables(self):
        tables = []
        try:
            # Try to import pandas for table handling
            import pandas as pd
            
            # Simple table detection using PyMuPDF
            for i, page in enumerate(self.pdf_doc, 1):
                # Look for potential tables by finding rectangles that might represent cells
                rect_areas = []
                for rect in page.search_for("   "):  # Search for multiple spaces as table indicator
                    rect_areas.append(rect)
                
                # Very simple table detection heuristic
                # Get all text blocks
                blocks = page.get_text("blocks")
                
                # Filter blocks that might be table cells (simple heuristic)
                potential_table_cells = []
                for block in blocks:
                    rect = fitz.Rect(block[:4])
                    text = block[4].strip()
                    if text and len(text) < 100:  # Likely a cell if short text
                        potential_table_cells.append((rect, text))
                
                # Skip if not enough potential cells
                if len(potential_table_cells) < 4:
                    continue
                
                # Simple row detection (group by similar y-coordinates)
                y_coords = {}
                for rect, text in potential_table_cells:
                    y_key = round(rect.y0 / 10) * 10  # Group by 10-pixel increments
                    if y_key not in y_coords:
                        y_coords[y_key] = []
                    y_coords[y_key].append((rect, text))
                
                # If we have at least 2 rows, consider it a table
                if len(y_coords) >= 2:
                    # Convert to table data format
                    table_data = []
                    for y_key in sorted(y_coords.keys()):
                        row = []
                        # Sort cells by x-coordinate
                        sorted_cells = sorted(y_coords[y_key], key=lambda c: c[0].x0)
                        for _, text in sorted_cells:
                            row.append(text)
                        if row:
                            table_data.append(row)
                    
                    # Add to tables list if we have data
                    if table_data and len(table_data) > 1:  # At least 2 rows
                        row_count = len(table_data)
                        col_count = max(len(row) for row in table_data)
                        tables.append({"page": i, "data": table_data, "size": f"{row_count}x{col_count}"})
        
        except Exception as e:
            print(f"Warning: Table extraction failed: {e}")
            # Minimal fallback for table detection
            for i, page in enumerate(self.pdf_doc, 1):
                text = page.get_text()
                # Check for table-like patterns
                if "\t" in text or "  " in text:
                    rows = text.split("\n")
                    table_data = [row.split("\t" if "\t" in row else "  ") for row in rows if row.strip()]
                    if table_data and sum(len(row) > 1 for row in table_data) > 3:  # At least 3 rows with multiple columns
                        row_count = len(table_data)
                        col_count = max(len(row) for row in table_data)
                        tables.append({"page": i, "data": table_data, "size": f"{row_count}x{col_count}"})
        
        return tables

# ========================= FILE EXPORT =========================
def save_text_and_metadata(extractor, output_dir="output", prefix=""):
    os.makedirs(output_dir, exist_ok=True)
    with open(f"{output_dir}/{prefix}_text.txt", "w", encoding="utf-8") as text_file:
        text_file.write(extractor.extract_text())

def save_tables_and_metadata(extractor, output_dir="output", prefix=""):
    os.makedirs(output_dir, exist_ok=True)
    with open(f"{output_dir}/{prefix}_tables.csv", "w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        for table in extractor.extract_tables():
            writer.writerow([f"Page {table['page']} - Size: {table['size']}"])
            writer.writerows(table["data"])
            writer.writerow([])

def save_links_and_metadata(extractor, output_dir="output", prefix=""):
    os.makedirs(output_dir, exist_ok=True)
    with open(f"{output_dir}/{prefix}_hyperlinks.txt", "w", encoding="utf-8") as link_file:
        for page, text, url in extractor.extract_hyperlinks():
            link_file.write(f"Page {page}: {text} -> {url}\n")

def save_images_metadata(extractor, output_dir="output", prefix=""):
    img_output_dir = f"{output_dir}/{prefix}_images"
    os.makedirs(output_dir, exist_ok=True)
    with open(f"{output_dir}/{prefix}_images_metadata.txt", "w", encoding="utf-8") as image_file:
        for page, path, fmt, resolution in extractor.extract_images(img_output_dir):
            image_file.write(f"Page {page}: {path} (Format: {fmt}, Resolution: {resolution})\n")

# ========================= MAIN =========================
if __name__ == "__main__":
    try:
        extractors = []
        
        # PPT Processing
        try:
            ppt_path = "project.pptx"
            ppt_loader = PPTLoader(ppt_path)
            ppt_extractor = PPTDataExtractor(ppt_loader)
            extractors.append(ppt_extractor)
            
            print("\n🔹 Extracting PPT Content...")
            save_text_and_metadata(ppt_extractor, prefix="ppt")
            save_tables_and_metadata(ppt_extractor, prefix="ppt")
            save_links_and_metadata(ppt_extractor, prefix="ppt")
            save_images_metadata(ppt_extractor, prefix="ppt")
            print("✅ PPT data successfully extracted!")
        except Exception as e:
            print(f"❌ Error processing PPT: {e}")
        
        # PDF Processing
        try:
            pdf_path = "report.pdf"
            pdf_loader = PDFLoader(pdf_path)
            pdf_extractor = PDFDataExtractor(pdf_loader)
            extractors.append(pdf_extractor)
            
            print("\n🔹 Extracting PDF Content...")
            save_text_and_metadata(pdf_extractor, prefix="pdf")
            save_tables_and_metadata(pdf_extractor, prefix="pdf")
            save_links_and_metadata(pdf_extractor, prefix="pdf")
            save_images_metadata(pdf_extractor, prefix="pdf")
            print("✅ PDF data successfully extracted!")
        except Exception as e:
            print(f"❌ Error processing PDF: {e}")
        
        # Store in MySQL (if we have any extractors)
        if extractors:
            print("\n🔹 Storing data in MySQL...")
            storage = MySQLStorage(extractors)
            storage.store_data()
            print("✅ Data successfully stored in MySQL!")
        
        print("\n✅ Processing completed!")
    except Exception as e:
        print(f"❌ Error in main process: {e}")