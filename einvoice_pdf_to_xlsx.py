# for PDF miner use
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed

# for regex use
import re

# for xlsx output
from datetime import datetime
import xlsxwriter

# for file system use
import os

# for argv use
import sys, getopt

# qr code and invoice code info refer to https://zhuanlan.zhihu.com/p/32315595

class ChineseAmount():
    chinese_amount_num = {
        '〇' : 0,
        '一' : 1,
        '二' : 2,
        '三' : 3,
        '四' : 4,
        '五' : 5,
        '六' : 6,
        '七' : 7,
        '八' : 8,
        '九' : 9,
        '零' : 0,
        '壹' : 1,
        '贰' : 2,
        '叁' : 3,
        '肆' : 4,
        '伍' : 5,
        '陆' : 6,
        '柒' : 7,
        '捌' : 8,
        '玖' : 9,
        '貮' : 2,
        #'两' : 2,
    }
    chinese_amount_unit = {
        '分' : 0.01,
        '角' : 0.1,
        '元' : 1,
        '圆' : 1,
        '十' : 10,
        '拾' : 10,
        '百' : 100,
        '佰' : 100,
        '千' : 1000,
        '仟' : 1000,
        '万' : 10000,
        '萬' : 10000,
        '亿' : 100000000,
        '億' : 100000000,
        '兆' : 1000000000000,
    }
    chinese_amount_exclude_char = {
        '整',
    }
    def convert_chinese_amount_to_number(self, chinese_amount):
        #chinese_amount = chinese_amount.strip('整') #remove unused char
        amount_number = 0
        for key, value in self.chinese_amount_unit.items():
            re_string = "(.{1})" + key
            regex = re.compile(re_string)
            result = re.search(regex, chinese_amount)
            if(result):
                if(result.group(1) in self.chinese_amount_num):
                    amount_number = amount_number + self.chinese_amount_num[result.group(1)] * value
        return amount_number


class eInvoicePDFParse():
    _debug_output_textfile = False
    def parse_pdf(self, pdf_path):
        with open(pdf_path, 'rb') as fp:
            parser = PDFParser(fp)
            doc = PDFDocument(parser)
            parser.set_document(doc)
            rsrcmgr = PDFResourceManager()
            laparams = LAParams(all_texts=True, boxes_flow=0.2, word_margin=0.5, detect_vertical=True)
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            extracted_text = ''
            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)
                layout = device.get_result()
                for lt_obj in layout:
                    if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                        #print(lt_obj)
                        #print(lt_obj.get_text())
                        extracted_text += lt_obj.get_text()
                    #else:
                        #print(lt_obj)
            return extracted_text
        
    einvoice_patten = {
        "invoice_code"      : r"\n([0-2][0-9]{11})\n", #发票代码 04403180xxx1
        "invoice_number"    : r"\n([0-9]{8})\n", #发票号码 1249xxx7
        "date"              : r"\n([0-9]{4})[^0-9/<>\+\-\*]+([0-9]{2})[^0-9/<>\+\-\*]+([0-9]{2}).*\n", 
        #"checksum"          : r"\n([0-9]{5} *[0-9]{5} *[0-9]{5} *[0-9]{5})\n", #校验码 12157 2xxx4 8xxx0 73378
        #"passcode"          : r"\n[0-9/<>\+\-\*]{28}\n", #密码区 03597///23xxx146/6xxx6>5>-/50 {0-9，"/<>+-*"}
        #"amount"            : r"\n[￥¥]+ *([0-9.]+)\n", #合计税额 < 合计金额 < 价税合计 ￥ 26.42
        "chinese_amount"    : r"\n([壹贰叁肆伍陆柒捌玖零亿万仟佰拾圆元角分整]+)\n",
        "itemName"          : r"\n(\*+[^0-9/<>\+\-\*]+\*+.+)\n", #*日用杂品*日用品
    }
    einvoice_result = [
        "invoice_code", "invoice_number",
        "date_year", "date_month", "date_day",
        "chinese_amount", "number_amount", "itemName",
        "file_name"
    ]

    def parse_einvoice_items(self, text):
        results = {}
        for key, patten in self.einvoice_patten.items():
            result = re.findall(patten, text)
            if(key == 'date'):
                results["date_year"] = result[0][0]
                results["date_month"] = result[0][1]
                results["date_day"] = result[0][2]
            else:
                results[key] = result[0]
            if(key == 'chinese_amount'):
                results["number_amount"] = ChineseAmount().convert_chinese_amount_to_number(result[0])
        return results

    def parse_einvoice_item_by_pdf(self, pdfFilePath):
        text = self.parse_pdf(pdfFilePath)
        if(self._debug_output_textfile):
            textFilePath = pdfFilePath.replace(".pdf", ".txt")
            with open(textFilePath, "w", encoding="utf-8") as f:
                f.write(text)
        einvoice_item_result = self.parse_einvoice_items(text)
        #einvoice_item_result["file_name"] = os.fsdecode(pdfFilePath)
        return einvoice_item_result


class eInvoicePDFtoExcel():
    def load_pdf_dir_get_einvoice_items(self, pdf_dir):
        eInvoicePDFParser = eInvoicePDFParse()
        pdfFiles = FileSystem().enumerate_pdf_in_folder(pdf_dir)
        eInvoice_results = []
        for pdfFile in pdfFiles:
            eInvoice_result = eInvoicePDFParser.parse_einvoice_item_by_pdf(pdfFile)
            eInvoice_results.append(eInvoice_result)
        return eInvoice_results

    einvoice_sheet_header = {
        "invoice_code"  : "发票代码",
        "invoice_number": "发票号码",
        "date"          : "开票日期",
        "chinese_amount": "价税合计 (大写)",
        "number_amount" : "价税合计 (小写)",
        "itemName"      : "货物或应税劳务、服务名称",
        #"file_name"     : "PDF文件名称",
    }
    def extract_items_to_xlsx(self, eInvoice_data, output_path):
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet("invoice")
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        row = 0
        col = 0
        # write header
        worksheet.write_string(row, col,     self.einvoice_sheet_header["invoice_code"])
        worksheet.write_string(row, col + 1, self.einvoice_sheet_header["invoice_number"])
        worksheet.write_string(row, col + 2, self.einvoice_sheet_header["date"])
        worksheet.write_string(row, col + 3, self.einvoice_sheet_header["chinese_amount"])
        worksheet.write_string(row, col + 4, self.einvoice_sheet_header["number_amount"])
        worksheet.write_string(row, col + 5, self.einvoice_sheet_header["itemName"])
        #worksheet.write_string(row, col + 6, self.einvoice_sheet_header["file_name"])
        row = 1
        for eInvoiceItem in eInvoice_data:
            dateStr = eInvoiceItem["date_year"] + '-' + eInvoiceItem["date_month"] + '-' + eInvoiceItem["date_day"]
            # Convert the date string into a datetime object.
            date = datetime.strptime(dateStr, "%Y-%m-%d")
            worksheet.write_string  (row, col,     eInvoiceItem["invoice_code"])
            worksheet.write_string  (row, col + 1, eInvoiceItem["invoice_number"])
            worksheet.write_datetime(row, col + 2, date, date_format)
            worksheet.write_string  (row, col + 3, eInvoiceItem["chinese_amount"])
            worksheet.write_number  (row, col + 4, eInvoiceItem["number_amount"])
            worksheet.write_string  (row, col + 5, eInvoiceItem["itemName"])
            #worksheet.write_string  (row, col + 6, eInvoiceItem["file_name"])
            row += 1
        workbook.close()

    def load_pdf_dir_output_xlsx(self, pdf_dir, output_xlsx):
        eInvoiceItems = self.load_pdf_dir_get_einvoice_items(pdf_dir)
        xlsxPath = os.path.join(pdf_dir, output_xlsx)
        self.extract_items_to_xlsx(eInvoiceItems, xlsxPath)

class FileSystem:
    def enumerate_pdf_in_folder(self, pdf_dir):
        pdfFileList = []
        directory = os.fsencode(pdf_dir)
        for file in os.listdir(directory):
            filename = os.fsdecode(file)
            if filename.endswith(".pdf"):
                pdfFilePath = os.path.join(pdf_dir, filename)
                pdfFileList.append(pdfFilePath)
        return pdfFileList


def main(argv):
    inputdir = None
    helpstr = 'Usage:\neinvoice_pdf_to_xlsx.py -d <pdf_dir>'
    try:
        opts, args = getopt.getopt(argv,"hd:",["pdf_dir="])
    except getopt.GetoptError:
        print(helpstr)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print(helpstr)
            sys.exit()
        elif opt in ("-d", "--pdf_dir"):
            inputdir = arg
    if not inputdir:
        print(helpstr)
        sys.exit(2)
    
    pdf2excel = eInvoicePDFtoExcel()
    pdf2excel.load_pdf_dir_output_xlsx(inputdir, "output.xlsx")

if __name__ == "__main__":
    main(sys.argv[1:])
