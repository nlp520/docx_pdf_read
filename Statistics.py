import os
import copy as copy
from urllib.request import urlopen
from pdfminer.pdfinterp import PDFResourceManager, process_pdf
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from io import StringIO
import docx
from win32com import client as wc

def doc_num(root_path):
    '''
    统计文档数量
    :param root_path:
    :return:
    '''
    count = 0
    if os.path.isdir(root_path):
        lis = os.listdir(root_path)
        for l in lis:
            count += doc_num(root_path + "/" + l)
        return count
    else:
        if "readme.txt" not in root_path:
            return 1
        return 0

def count_words(root_path):
    count = 0
    if os.path.isdir(root_path):
        lis = os.listdir(root_path)
        for l in lis:
            count += count_words(root_path + "/" + l)
        return count
    else:
        if ".docx" in root_path:
            content = clear(read_docx(root_path))
            number = content.count(".")
            print(root_path)
            print("句子数量为：", number)
            return number
        # elif ".doc" in root_path:
        #     root_path = save_as_docx(root_path)
        #     content = clear(read_docx(root_path))
        #     number = len(content.split(" "))
        #     print(root_path)
        #     print("单词数量为：", number)
        #     return number
        elif ".pdf" in root_path:
            try:
                content = clear(readPDF(root_path))
                number = content.count(".")
                print(root_path)
                print("句子数量为：", number)
                return number
            except:
                return 0
        else:
            return 0

def readPDF(pdfFile):
    '''
    读取pdf
    :param pdfFile:
    :return:
    '''
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)

    process_pdf(rsrcmgr, device, open(pdfFile, "rb"))
    device.close()

    content = retstr.getvalue()
    retstr.close()
    content = content.replace("\n", "")
    return content

def read_docx(path):
    '''
    读取docx中的内容
    :param path:
    :return:
    '''
    file = docx.Document(path)
    context = ""
    for para in file.paragraphs:
        context += para.text
    context = context.replace("\n", "")
    return context

def save_as_docx(path):
    word = wc.Dispatch('Word.Application')
    print(path)
    doc = word.Documents.Open(path)  # 目标路径下的文件
    doc.SaveAs(path+"x", 12, False, "", True, "", False, False, False,
               False)  # 转化后路径下的文件
    doc.Close()
    word.Quit()
    return path+"x"

def clear(content):
    import re
    content = re.sub("[^a-zA-Z\s?,.]", "", content)
    return content

result = doc_num("G:/sui/翻译语料/专业语料/专业语料")
print("英文文档的数量大约为：",result//2)
word_num = count_words("G:/sui/翻译语料/专业语料/专业语料")
print("全部的句子的数量大约为：", word_num)

'''
英文文档的数量大约为： 224
全部的英文单词的数量大约为： 2012514
全部的句子的数量大约为： 208887
'''
