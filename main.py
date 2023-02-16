import time
from docx import Document
from win32com import client as wc
import os
import re
import shutil


def docx_remove_content(doc, finish_doc):
    # 定义需要去除的内容
    content_to_remove = '''【提示】本文档来源自第一范文网（https://www.diyifanwen.com），第一范文网是中国最大的范文网站，专注于提供各种优质工作学习办公文档，欢迎访问。
微信搜索公众号“第一范文网”，关注后可方便查找下载各种文档。
转发文档请保留以上标记，谢谢！'''

    # 打开doc文件
    doc = Document(doc)

    # 遍历doc文件中的段落
    for para in doc.paragraphs:
        # 如果段落中包含需要去除的内容，使用正则表达式替换为空字符串
        if re.search(content_to_remove, para.text):
            para.text = re.sub(content_to_remove, '', para.text)

    # 遍历doc文件中的表格
    # for table in doc.tables:
    #     # 遍历表格中的行
    #     for row in table.rows:
    #         # 遍历行中的单元格
    #         for cell in row.cells:
    #             # 如果单元格中包含需要去除的内容，使用正则表达式替换为空字符串
    #             if re.search(content_to_remove, cell.text):
    #                 cell.text = re.sub(content_to_remove, '', cell.text)

    # 保存修改后的doc文件
    doc.save(finish_doc)


def get_word_pages(in_file):
    pages = 1
    try:
        word = wc.Dispatch("Word.Application")
        try:
            doc = word.Documents.Open(in_file)
            word.ActiveDocument.Repaginate()
            pages = word.ActiveDocument.ComputeStatistics(2)
            doc.Close()
            word.Quit()
            return pages
        except Exception as e:
            print(e)
        finally:
            return pages
    except Exception as e:
        print(e)
    finally:
        return pages


def doc2docx(in_file, out_file):
    try:
        word = wc.Dispatch("Word.Application")
        try:
            print(in_file)
            print(out_file)
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, 12, False, "", True, "", False, False, False, False)
            print('转换成功')
            doc.Close()
            word.Quit()
        except Exception as e:
            print(1111)
            print(e)
    except Exception as e:
        print(e)
    exit(1)


word_dir = "G:\\www.diyifanwen.com\\导游词\\北京导游词"
finish_dir = "G:\\www.diyifanwen.com\\导游词\\北京导游词_finish"
doc2docx_dir = "G:\\www.diyifanwen.com\\导游词\\北京导游词_doc2docx"

if __name__ == '__main__':
    if not os.path.exists(finish_dir):
        os.mkdir(finish_dir)

    if not os.path.exists(doc2docx_dir):
        os.mkdir(doc2docx_dir)

    files = sorted(os.listdir(word_dir))
    for file in files:
        if os.path.splitext(file)[1] in [".doc", ".docx"]:
            print(file)
            if os.path.splitext(file)[1] == ".docx":
                # 将文件复制到doc2docx_dir目录
                print("复制文件")
                shutil.copyfile(word_dir + "\\" + file, doc2docx_dir + "\\" + file)
            elif os.path.splitext(file)[1] == ".doc":
                # 将doc文件转化为docx文件
                print("转化文件")
                doc2docx(word_dir + "\\" + file, doc2docx_dir + "\\" + os.path.splitext(file)[0] + ".docx")
                time.sleep(3)
            # 去除word页眉和页脚
            doc2docx_file = doc2docx_dir + "\\" + os.path.splitext(file)[0] + ".docx"
            finish_doc = finish_dir + "\\" + os.path.splitext(file)[0] + ".docx"

            word_pages = get_word_pages(doc2docx_file)
            if 3 <= word_pages <= 60:
                docx_remove_content(doc2docx_file, finish_doc)
