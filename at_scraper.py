import os
import pandas as pd
import fitz
import re
from rich import print

class ATPDFScraper:
    """Web scraper for extracting product details from PDFs."""
    def __init__(self, excel_path: str, file1: str, file2: str, output_filename: str):
        self.filepath = excel_path
        self.file1 = file1
        self.file2 = file2
        self.output_filename = output_filename
        self.df = pd.read_excel(self.filepath, sheet_name="Master")
    
    def extract_procuity_data(self, pdf_path):
        """Extract product details from the ProCuity PDF."""
        print(f"[cyan]Extracting ProCuity data from:[/cyan] {pdf_path}")
        doc = fitz.open(pdf_path)
        pdf_text = "\n".join([page.get_text("text") for page in doc])
        doc.close()

        try:
            for page in doc:
                text = page.get_text("text")  # Extract text from page
                lines = text.split("\n")  # Split text into lines

                extracted_text = []
                table_data = []
                is_table = False
            start_description = pdf_text.split("Brilliance in a bed")[1]
            description_text = start_description.split("L model")[0].strip()
        except IndexError:
            description_text = ""  # In case "Brilliance in a bed" or "L model" isn't found

        
        # Extract product details using regex
        model_number = re.search(r"Model number\s*(\d+)", pdf_text)
        width = re.search(r"Width\s*(\d+.*?cm)", pdf_text)
        height = re.search(r"Height range.*?High:\s*(\d+.*?cm).*?Low:\s*(\d+.*?cm)", pdf_text, re.DOTALL)
        volts = re.search(r"Volts:\s*(\d+-\d+ VAC)", pdf_text)
        amps = re.search(r"Ampere rating:\s*(\d+A)", pdf_text)
        hertz = re.search(r"Frequency:\s*(\d+/\d+ Hz)", pdf_text)
        plug_type = re.search(r"Hospital grade plug:\s*(\d+ VAC, \d+ Hz, \d+A)", pdf_text)
        weight = re.search(r"Safe working load\s*(\d+.*?kg)", pdf_text)
        description_match = re.search(r"Brilliance in a bed(.*?)helping hospitals standardize their bed fleet and improve caregiver efficiencies.", pdf_text, re.DOTALL)
        product_description = description_match.group(1).strip() if description_match else ""
        # Assign extracted values
        return {
            "mfr website": "https://www.stryker.com",
            "mfr name": "Stryker",
            "model name": "ProCuity",
            "mfr number": model_number.group(1) if model_number else "",
            "product description": description_text,
            "amps": amps.group(1) if amps else "",
            "volts": volts.group(1) if volts else "",
            "hertz": hertz.group(1) if hertz else "",
            "plug_type": plug_type.group(1) if plug_type else "",
            "weight": weight.group(1) if weight else "",
            "height": f"High: {height.group(1)}, Low: {height.group(2)}" if height else "",
            "width": width.group(1) if width else "",
            "Specification Sheet (pdf)": os.path.basename(pdf_path),
            "Product URL": "https://www.stryker.com"
        }
    
    def extract_sterilgard_data(self, pdf_path):
        """Extract product details from the SterilGARD SGX04 PDF."""
        print(f"[cyan]Extracting SterilGARD data from:[/cyan] {pdf_path}")
        doc = fitz.open(pdf_path)
        pdf_text = "\n".join([page.get_text("text") for page in doc])
        doc.close()

        # Extract product details using regex
        model_number = re.search(r"MODEL\s*(SG\d+)", pdf_text)
        width = re.search(r"Nominal Size\s*(\d+.*?Meters)", pdf_text)
        height = re.search(r"Cabinet Height\s*(\d+.*?mm)", pdf_text)
        volts = re.search(r"Service Requirements.*?(\d+ V AC)", pdf_text, re.DOTALL)
        amps = re.search(r"(\d+ A), 50/60 Hz", pdf_text)
        hertz = "50/60 Hz"
        plug_type = "Listed plug for destination country"
        weight = re.search(r"Weight.*?(\d+ Kg)", pdf_text)

        # Assign extracted values
        return {
            "mfr website": "https://www.bakerco.com",
            "mfr name": "Baker",
            "model name": "SterilGARD SGX04",
            "mfr number": model_number.group(1) if model_number else "",
            "product description": "SterilGARD SGX04 Class II, Type A2 Biosafety Cabinet.",
            "amps": amps.group(1) if amps else "",
            "volts": volts.group(1) if volts else "",
            "hertz": hertz,
            "plug_type": plug_type,
            "weight": weight.group(1) if weight else "",
            "height": height.group(1) if height else "",
            "width": width.group(1) if width else "",
            "Specification Sheet (pdf)": os.path.basename(pdf_path),
            "Product URL": "https://www.bakerco.com"
        }

    def run(self):
        """Main function to extract product details and save them to an Excel file."""
        pdf_files = [(self.file1, self.extract_procuity_data), (self.file2, self.extract_sterilgard_data)]
        new_rows = []
        
        for pdf, extractor in pdf_files:
            pdf_path = os.path.join(os.getcwd(), pdf)
            if os.path.exists(pdf_path):
                product_data = extractor(pdf_path)
                new_rows.append(product_data)
            else:
                print(f"[red]File not found:[/red] {pdf_path}")

        # Append new rows to DataFrame
        new_df = pd.DataFrame(new_rows)
        self.df = pd.concat([self.df, new_df], ignore_index=True)
        self.df.to_excel(self.output_filename, index=False, sheet_name="Master")
        print("[green]Data extraction complete. File saved![/green]")


# ---------------------------------------- RUN THE CODE ----------------------------------------
if __name__ == "__main__":
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    scraper = ATPDFScraper(
        excel_path="AT -WebScrape Content Template (Master).xlsx",
        file1="2020 ProCuity Spec Sheet JB Mkt Lit 2077 07 OCT 2020 REV C 1.pdf",
        file2="SterilGARD-SGX04-Product-Specifications-RevE.pdf",
        output_filename="output/AT-Scrape-output.xlsx",
    )
    scraper.run()
