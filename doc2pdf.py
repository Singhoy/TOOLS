# -*- encoding: utf-8 -*-
import os

from win32com import client


def doc2pdf(doc_name, pdf_name, doc_path):
    """
    word文件转pdf
    :param doc_path: word文档所在路径
    :param doc_name word文件名称
    :param pdf_name 转换后pdf文件名称
    """
    old_pdf = pdf_name
    doc_name = doc_path + '\\' + doc_name
    pdf_name = doc_path + '\\' + pdf_name
    try:
        print('***格式转换中，请稍候。。。***')
        word = client.DispatchEx("Word.Application")
        if os.path.exists(pdf_name):
            os.remove(pdf_name)
        worddoc = word.Documents.Open(doc_name, ReadOnly=1)
        worddoc.SaveAs(pdf_name, FileFormat=17)
        worddoc.Close()
        print('***格式转换成功！***')
        print(f'新文件{old_pdf}在{doc_path}')
        return pdf_name
    except Exception as e:
        print('***转换出错，程序中断***')
        print(e)
        return 1
      
      
if __name__ == '__main__':
    dc_name = input('请输入doc文件名(带上后缀名)：')
    path = input('请输入文件所在路径：')
    pf_name = input('请输入转换后的文件名(带上后缀名)：')
    doc2pdf(dc_name, pf_name, path)
