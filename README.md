# Automatic Certificate of Analysis with Product Specifications Maker

This project automates the creation of Certificates of Analysis (CoA) using Excel, Power Query, VBA, and Excel functions like VLOOKUP. It simplifies the process by integrating data from multiple sources and generating a product's certificate of analysis with minimal manual effort.

## Features

- **Data Integration**: 
  - Inputs are gathered from:
    1. A shareable Google Sheets file where the team can input laboratory results and product standards for each product.
    2. A Product Information workbook containing product description, uses, dosage, storage, shelf life, ingredients list, and packing details.
  
- **Power Query**: 
  - Consolidates and transforms the raw data into a final sheet in Excel.

- **VBA Automation**: 
  - Automates the creation of the Certificate of Analysis, allowing users to generate the certificate with just 2-3 clicks, significantly improving efficiency.

## How It Works

1. **Input Data**: 
   - Lab results and standards are entered in the Google Sheets file, where each member may input the physical/chemical analysis results.
   - Product information is loaded from the workbook.
   
2. **Data Processing**: 
   - Power Query loads the necessary data into a new Excel sheet, preparing it for the CoA.
   
3. **Certificate Generation**: 
   - With a simple VBA macro, the Certificate of Analysis is automatically generated, reducing manual effort and the risk of errors.

## Benefits

- Automates the process of creating a Certificate of Analysis.
- Reduces manual data entry and improves accuracy.
- Saves time by allowing the team to generate CoAs with just a few clicks.
