# encoding=utf-8

import comtypes.client
import os
import logging
import configparser


# PDFFileFormat=17
# PPTFileFormat=32


def app():
    filenames = os.listdir(office_folder)
    for item in filenames:
        origin_path = os.path.join(office_folder, item)
        if item.index(".") > -1:
            pdf_filename = item[0: item.rindex(".")] + ".pdf"
        else:
            pdf_filename = item + ".pfd"
        pdf_path = os.path.join(pdf_folder, pdf_filename)
        if item.endswith(".doc") or item.endswith(".docx"):
            c_word2pdf(origin_path, pdf_path)
        elif item.endswith(".xls") or item.endswith(".xlsx"):
            c_xls2pdf(origin_path, pdf_path)
        elif item.endswith(".ppt") or item.endswith(".pptx"):
            c_ppt2pdf(origin_path, pdf_path)
    print("complete")

def c_word2pdf(origin_path, pdf_path):
    logger = logging.getLogger("c_word2pdf")
    try:
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(origin_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        logger.info("文档转化成功，【%s】", origin_path)
    except:
        logger.info("文档转化失败，【%s】", origin_path)


def c_ppt2pdf(origin_path, pdf_path):
    logger = logging.getLogger("c_ppt2pdf")
    try:
        application = comtypes.client.CreateObject('Powerpoint.Application')
        application.Visible = 1
        office = application.Presentations.Open(origin_path)
        office.SaveAs(pdf_path, FileFormat=32)
        office.Close()
        application.Quit()
        logger.info("文档转化成功，【%s】", origin_path)
    except:
        logger.info("文档转化失败，【%s】", origin_path)


def c_xls2pdf(origin_path, pdf_path):
    logger = logging.getLogger("c_xls2pdf")
    try:
        application = comtypes.client.CreateObject('Excel.Application')
        application.Visible = 1
        office=application.Workbooks.Open(origin_path)
        office.ExportAsFixedFormat(0, pdf_path)
        office.Close()
        application.Quit()
        logger.info("文档转化成功，【%s】", origin_path)
    except:
        logger.info("文档转化失败，【%s】", origin_path)


if __name__ == "__main__":
    logging.basicConfig(filename="logger.log", level=logging.INFO, format='%(name)s - %(levelname)s - %(asctime)s - %(message)s')
    cp = configparser.ConfigParser()
    cp.read("config.ini", encoding="utf-8")
    cp.options("config")
    office_folder = cp.get("config", "office_folder")
    pdf_folder = cp.get("config", "pdf_folder")
    app()
