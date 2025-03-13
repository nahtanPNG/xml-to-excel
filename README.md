# 📄 XML NFe to Excel

A Python project that reads **XML files from Electronic Invoices (NFe)**, extracts product information, and generates an organized **Excel (.xlsx)** table.

## 📊 Features

- 📂 Reads all `.xml` files from a directory.
- 📄 Extracts the following information from each NFe:
    - Emission Date
    - Product
    - Quantity
    - Unit Value
    - Total Value
- 📊 Creates a consolidated Excel file with a separate sheet for each NFe.
- ✅ Displays the **Emission Date** only in the first row of each sheet.

## 📋 Requirements

- Python 3.10 or higher
- Required libraries:

```bash
pip install pandas openpyxl
```

## 📁 Project Structure

```
xml_to_excel/
├── main.py          # Main script
├── requirements.txt # Dependencies
└── README.md        # Project documentation
```

## 🚀 How to Run

1. **Clone the repository**

```bash
git clone https://github.com/your-username/xml-to-excel.git
cd xml-to-excel
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **Run the script**

```bash
python main.py <your/xml/files/directory>
```

4. **Output:**

- The `output.xlsx` file will be generated in the specified directory, with a sheet for each processed XML file.

## 📊 Example Output

| Date                | Product      | Quantity | Unit Value | Total Value |
|---------------------|--------------|----------|------------|-------------|
| 2024-02-07T14:30:00 | Product A    | 2        | 50.00      | 100.00      |
|                     | Product B    | 1        | 30.00      | 30.00       |
|                     | Product C    | 5        | 10.00      | 50.00       |

## 🛠️ Customization

If you need to adjust the extracted fields, edit the `find` calls inside the loop that processes products in the XML.

## 📌 Notes

- Ensure the XML files follow the NFe standard and contain the correct namespaces.
- Each sheet name in the Excel file is based on the XML file name (limited to 31 characters).

## 📄 License

This project is licensed under the [MIT License](LICENSE).

