# **Open PO Report Generator - README**

## **Overview**

This repository contains a Python script designed to automate the generation of individual purchase order (PO) reports for suppliers. The code processes a consolidated ERP PO report, filters and formats the data, and creates individual reports for each supplier. These reports are saved in an organized directory structure, ensuring that each supplier's report is professionally formatted and easy to access. This automation helps streamline procurement processes and digitizes supply chain management by reducing manual work and increasing efficiency.

Additionally, this script can be integrated with **AI-based LLM (Large Language Model) bots** to enhance automation in supply chain operations, providing real-time data and generating reports through conversational interfaces.

---

## **Key Features**

- **Automated Report Generation**: Automatically creates individual supplier reports from a consolidated ERP PO report.
- **Data Preprocessing**: Filters out irrelevant POs (e.g., completed ones) and ensures data is sorted and structured correctly.
- **Professional Excel Formatting**: Applies consistent formatting to Excel files, including header styling, border formatting, and column width adjustments.
- **Scalability**: Handles large datasets and can easily be adapted for growing supplier networks.
- **Integration with LLM bots**: Can be used as part of a larger AI-based supply chain system where bots assist with report generation and data extraction.

---

## **Prerequisites**

Before running the script, ensure that you have the following Python libraries installed:

- **pandas**: A powerful data manipulation and analysis library.
- **openpyxl**: A library for reading and writing Excel (xlsx) files.

To install the required libraries, you can use `pip`:

```bash
pip install pandas openpyxl
```

---

## **Getting Started**

### **1. Clone the Repository**

Start by cloning this repository to your local machine:

```bash
git clone https://github.com/your-username/po-report-generator.git
```

### **2. Input Data Format**

This script expects the input data to be in an Excel file with the following columns:

- **Po Creation Date**: Date when the purchase order was created.
- **PO Qty Due**: Quantity of the PO that is still due.
- **Supplier Name**: The name of the supplier.

Ensure that your input file is structured similarly to the example format.

### **3. Input File Path**

Place your ERP PO report (Excel file) in the `Input Data` folder within the repository. The file should be named `PO_Report.xlsx`, or you can update the script to reflect the correct path and file name.

```python
input_file_path = 'Input Data/PO_Report.xlsx'
```

---

## **Running the Script**

To run the script, execute the following command:

```bash
python generate_po_reports.py
```

This will:

1. Load the input Excel file.
2. Process and filter the data (remove completed POs).
3. Generate individual reports for each supplier.
4. Apply professional formatting to the Excel files.
5. Save the reports in a timestamped directory within the `Output Data` folder.

For example, if the script is run on January 19, 2025, the output folder structure will look like this:

```
Output Data/
    Supplier_PO_Reports/
        PO_Reports_20250119_123456/
            Supplier_A_PO_Report_20250119_123456.xlsx
            Supplier_B_PO_Report_20250119_123456.xlsx
            ...
```

Each report will be saved as a separate Excel file named after the supplier.

---

## **Detailed Explanation of Code**

The script is divided into multiple sections, each handling a different aspect of the process:

### **1. Read and Process Input Data**

The script reads the ERP PO report into a pandas DataFrame, which allows for easy manipulation and filtering of the data.

### **2. Data Preprocessing**

- **Date Conversion**: Converts the "Po Creation Date" column into a `datetime` format to ensure proper sorting.
- **PO Filtering**: Removes POs with a quantity due of zero, which indicates they are completed.
- **Sorting**: Sorts the data based on the PO Creation Date in ascending order.

### **3. Output Directory Setup**

The script creates an organized directory structure for storing the reports. Each run creates a new timestamped subfolder to ensure that reports from different runs are kept separate.

### **4. Excel Formatting**

The script uses `openpyxl` to apply professional formatting to the generated Excel reports. This includes:

- **Header Formatting**: Bold, centered headers with a dark blue background.
- **Cell Formatting**: Borders around all cells and appropriate text alignment (right-align for numbers and left-align for text).
- **Auto-adjust Column Width**: Columns are auto-adjusted to fit the content, with a maximum width of 50 characters.
- **Freeze Header Row**: The header row is frozen to remain visible as users scroll through the data.

### **5. Generate Individual Supplier Reports**

The script processes each unique supplier in the dataset, filters the data accordingly, and creates an individual Excel report for each one. Each report is formatted and saved to the output folder.

---

## **Customization Options**

You can customize the script by modifying the following:

- **Input File Path**: Change the path to your input file.
- **Header Fill Color**: Adjust the header color by changing the `start_color` and `end_color` in the `header_fill` variable.
- **Font Sizes**: Modify the font size for headers and data by changing the `Font()` objects.
- **Column Width**: Adjust the maximum allowed column width (currently set to 50).

---

## **Use Cases in Supply Chain Digitization**

This script is an excellent starting point for **digitizing supply chain management**, especially in procurement processes. Automating the generation of supplier PO reports allows supply chain managers to focus on more strategic tasks rather than manual report generation. The solution can be scaled up to handle large datasets, making it ideal for businesses with many suppliers.

### **Integration with AI and LLM Bots**

The script can be integrated into **AI-based systems** and **LLM bots** for further automation. For example, an LLM bot could use natural language commands to trigger the report generation process, making it more interactive and user-friendly. This integration could enable real-time insights and data retrieval, allowing supply chain managers to make faster, data-driven decisions.

---

## **Conclusion**

The **PO Report Generator** is a simple yet powerful tool to automate and streamline PO report generation, making supply chain management more efficient and data-driven. By integrating this with AI-powered bots and automation systems, businesses can take their supply chain digitization to the next level, ensuring accuracy, speed, and scalability.

---

## **License**

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## **Contributing**

Feel free to fork the repository and submit pull requests if you have improvements or bug fixes. Contributions are always welcome!

---

## **Contact**

For any questions or support, please contact [srekatpure@gmail.com].
