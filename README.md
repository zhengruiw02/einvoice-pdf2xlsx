# einvoice PDF to xlsx

# Features
- Read PDFs from given folder and extract items output to xlsx sheet
- Extract items include :
    - 发票代码
    - 发票号码
    - 开票日期
    - 价税合计 (大写)
    - 价税合计 (小写)
    - 货物或应税劳务、服务名称
# Dependence
- python 3.7
- pdfminer.six
- xlsxwriter

# Quick Start
Use following command in cli with parameter ```<your_dir_with_einvoice_pdf>``` for a directory which contains your einvoice PDF files.
```bash
python einvoice_pdf_to_xlsx.py -d <your_dir_with_einvoice_pdf>
```
Then you will get ```output.xlsx``` in your given directory.
