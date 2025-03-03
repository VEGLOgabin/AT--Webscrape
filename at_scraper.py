import os
import pandas as pd
import fitz
import re
from rich import print
import pdfplumber

class ATPDFScraper:
    """Web scraper for extracting product details from PDFs."""
    def __init__(self, excel_path: str, file1: str, file2: str):
        self.filepath = excel_path
        self.file1 = file1
        self.file2 = file2
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
        data =  {
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
    
        # Append new rows to DataFrame
        new_df = pd.DataFrame([data])
        self.df = pd.concat([self.df, new_df], ignore_index=True)
        self.df.to_excel(self.output_filename, index=False, sheet_name="Master")
        print("[green]Data extraction complete. File saved![/green]")
    
    def extract_sterilgard_data(self, pdf_path):
        # Load the PDF
        pdf_path = "SterilGARD-SGX04-Product-Specifications-RevE.pdf"
        pdf = pdfplumber.open(pdf_path)

        # Initialize a dictionary to store extracted data
        data = {
            'mfr website': '',
            'mfr name': 'Baker',
            'model name': '',
            'mfr number': '',
            'unit cost': '',
            'product description': '',
            'amps': '',
            'volts': '',
            'watts': '',
            'phase': '',
            'hertz': '',
            'plug_type': '',
            'emergency_power Required (Y/N)': 'N',
            'dedicated_circuit Required (Y/N)': 'Y',
            'tech_conect Required': '',
            'btu ': '',
            'dissipation_type': 'Air',
            'water_cold Required (Y/N)': 'N',
            'water_hot  Required (Y/N)': 'N',
            'drain Required (Y/N)': 'N',
            'water_treated (Y/N)': 'N',
            'steam  Required(Y/N)': 'N',
            'vent  Required (Y/N)': 'Y',
            'vacuum Required (Y/N)': 'N',
            'ship_weight': '',
            'weight': '',
            'depth': '',
            'height': '',
            'width': '',
            'ada compliant (Y/N)': 'N',
            'green certification? (Y/N)': 'N',
            'antimicrobial coating (Y/N)': 'N',
            'Specification Sheet (pdf)': pdf_path,
            'Brochure (pdf)': '',
            'Manual/IFU (pdf)': '',
            'Product URL': '',
            'CAD (dwg)': '',
            'REVIT (rfa)': '',
            'Seismic document': '',
            'Product Image (jpg)': '',
            'Notes': ''
        }
        # Initialize a list to store data for all models
        prod1 = data.copy()
        prod2 = data.copy()
        prod3 = data.copy()

        # print(len(pdf.pages))
        pages_length = len(pdf.pages)
        if pages_length == 9:
            pages = pdf.pages
            page1 = pages[0]
            if page1:
                text1 = page1.extract_text()
                page1_text = []
                lines1 = text1.split("\n")  
                for line in lines1:
                        line = line.strip()
                        page1_text.append(line)
                # print(page1_text)
                prod_descrip = page1_text[2] + " " + page1_text[3]
                prod1["product description"] = prod_descrip
                prod2["product description"] = prod_descrip
                prod3['product description'] = prod_descrip

            page2 = pages[1]
            if page2:
                page2_table = page2.extract_table()
                # print(page2_table)
                prod1['mfr number'] = page2_table[0][2]
                prod2['mfr number'] = page2_table[0][4]
                prod3['mfr number'] = page2_table[0][5]

                # Extract dimensions and weights for SG404
                prod1["width"] = int(page2_table[3][2].split("[")[-1].strip("[]").split("x")[0].replace(",",""))* 0.0393701
                prod1["depth"] = int(page2_table[3][2].split("[")[-1].strip("[]").split("x")[1].replace(",","").replace("mm", ""))* 0.0393701  
                prod1["height"] = int(page2_table[4][2].split("[")[1].split("]")[0].strip("[]").replace(",","").replace("mm", ""))* 0.0393701
                prod1["weight"] = page2_table[7][2].split("lbs")[0].strip() 
                prod1["ship_weight"] = page2_table[12][2].split("lbs")[0].strip() 
                # print(prod1["ship_weight"]) 

                # Extract dimensions and weights for SG504
                prod2["width"] = int(page2_table[3][4].split("[")[-1].strip("[]").split("x")[0].replace(",",""))* 0.0393701 
                prod2["depth"] = int(page2_table[3][4].split("[")[-1].strip("[]").split("x")[1].replace(",","").replace("mm", ""))* 0.0393701    
                prod2["height"] = int(page2_table[4][4].split("[")[1].split("]")[0].strip("[]").replace(",","").replace("mm", ""))* 0.0393701 
                prod2["weight"] = page2_table[7][4].split("lbs")[0].strip()   
                prod2["ship_weight"] = page2_table[12][4].split("lbs")[0].strip() 

                # Extract dimensions and weights for SG604
                prod3["width"] = int(page2_table[3][5].split("[")[-1].strip("[]").split("x")[0].replace(",",""))* 0.0393701
                prod3["depth"] = int(page2_table[3][5].split("[")[-1].strip("[]").split("x")[1].replace(",","").replace("mm", ""))* 0.0393701      
                prod3["height"] = int(page2_table[4][5].split("[")[1].split("]")[0].strip("[]").replace(",","").replace("mm", ""))* 0.0393701  
                prod3["weight"] = page2_table[7][5].split("lbs")[0].strip()   
                prod3["ship_weight"] = page2_table[12][5].split("lbs")[0].strip()  

            page4 = pages[3]
            if page4:
                page4_table = page4.extract_table()
                volt_amps_hetz_phase = page4_table[2][2]
                volt_amps_hetz_phase = volt_amps_hetz_phase.split(",")
                volts = volt_amps_hetz_phase[0].split("V")[0]
                amps = volt_amps_hetz_phase[1].replace("A","")
                hetz = volt_amps_hetz_phase[2].replace("Hz", "")
                phase = volt_amps_hetz_phase[3]
                prod1['volts'] = volts
                prod2["volts"] = volts
                prod3["volts"] = volts

                prod1["amps"] =  amps
                prod2["amps"] =  amps
                prod3["amps"] =  amps

                prod1["hertz"] = hetz
                prod2["hertz"] = hetz
                prod3["hertz"] = hetz

                prod1['phase'] = phase
                prod2['phase'] = phase
                prod3['phase'] = phase
                

            page5 = pages[4]
            if page4:
                page5_table = page5.extract_table()
                # print(page5_table)
                prod1["btu "] = str(page5_table[22][1]).replace("Btu/Hr", "")
                prod2["btu "] = str(page5_table[22][3]).replace("Btu/Hr", "")
                prod3["btu "] = str(page5_table[22][4]).replace("Btu/Hr", "")
            

        all_data = [prod1, prod2, prod3]
        # # Create a DataFrame for all models
        df = pd.DataFrame(all_data)

        # Save to Excel
        df.to_excel("output/SterilGARD_SGX04_All_Models_Extracted_Data.xlsx", index=False)

        # print("Data for all models extracted and saved to Excel.")

    def run(self):
        """Main function to extract product details and save them to an Excel file."""
        pdf_files = [(self.file1, self.extract_procuity_data), (self.file2, self.extract_sterilgard_data)]
        new_rows = []
        
        for pdf, extractor in pdf_files:
            pdf_path = os.path.join(os.getcwd(), pdf)
            if os.path.exists(pdf_path):
                extractor(pdf_path)
                
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
    )
    scraper.run()
