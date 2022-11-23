import os
import win32api
import win32con
import tempfile
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfFileReader, PdfFileWriter
from natsort import ns, natsorted
from tempfile import TemporaryDirectory
from PIL import Image

Image.MAX_IMAGE_PIXELS = None


def TkPath():
    '''
    获取初始选定的文件夹路径
    :return: 根目录
    '''
    root = tk.Tk()
    root.withdraw()
    path_root = filedialog.askdirectory()
    return path_root


def Ch2PDF(pic, path, num):
    '''
    图片转PDF，如果图片过大则切割图片后保存为多张PDF
    :param pic: 图片 （含完整路径）
    :param path: 文件路径
    :param num: 计数器，用来给文件命名
    '''
    img = Image.open(pic)
    img_h = img.size[1]
    img_w = img.size[0]
    if (img_h // 65500) > 0:
        s = (img_h // 65500) + 1
        for page in range(s):
            region = img.crop((0, page * img_h // s, img_w, (page + 1) * img_h // s))
            region.save(path + '\\' + str(num) + '-' + str(page) + '.pdf', 'pdf')
            pass
    else:
        img.save(path + '\\' + str(num) + '.pdf', 'pdf')
        pass
    pass


def p2p(pdf, path, num):
    '''
    存放PDF至临时缓存
    :param pdf: 需要加载的pdf
    :param path: 文件路径
    :param num: 计数器
    :return:
    '''
    pdf_save = PdfFileReader(pdf)
    pdfWriter = PdfFileWriter()
    pdf_num = 0

    for page in range(pdf_save.numPages):
        pdfWriter.addPage(pdf_save.getPage(page))
        print('正在合并第{}页PDF文件'.format(pdf_num))
        pdf_num += 1
        pass
    with open(path + '\\' + str(num) + '.pdf', 'wb') as f:
        pdfWriter.write(f)
        pass
    pass


def main():
    path_root = TkPath()

    # 遍历文件夹获取所有文件路径
    file_list = []
    for dir_cur, dirs, files in os.walk(path_root, topdown=False):
        if files:
            for file in files:
                path_cur = dir_cur + '\\' + file
                file_list.append(path_cur)
                pass
            pass
        pass

    file_list = natsorted(file_list, alg=ns.REAL | ns.LOCALE | ns.IGNORECASE)
    print('共遍历到{}个文件'.format(len(file_list)))

    page_num = 0
    with tempfile.TemporaryDirectory() as tmpdir:
        for pic in file_list:
            if pic[-3:] == 'pdf':
                print('正在转置第{}个文件，该文件是PDF文件，转置处理中'.format(page_num), pic)
                p2p(pic, tmpdir, page_num)
                page_num += 1
            else:
                print('正在转置第{}个文件'.format(page_num), pic)
                Ch2PDF(pic, tmpdir, page_num)
                page_num += 1
                pass
            pass

        pdf_list = []
        for dir_cur, dirs, files in os.walk(tmpdir):
            if files:
                for file in files:
                    path_cur = dir_cur + '\\' + file
                    pdf_list.append(path_cur)
                    pass
                pass
            pass

        # for i in pdf_list:
        #     print(i)
        # print('############')

        pdf_list = natsorted(pdf_list, alg=ns.REAL | ns.LOCALE | ns.IGNORECASE)

        # for i in pdf_list:
        #     print(i)

        pdf_page = 0

        if len(pdf_list) > 1:
            # pdf_final = PdfFileReader(pdf_list[0])
            pdf_finalWriter = PdfFileWriter()
            for pdf_final in pdf_list:
                pdf2pdf = PdfFileReader(pdf_final)
                if pdf2pdf.numPages > 1:
                    for page in range(pdf2pdf.numPages):
                        pdf_finalWriter.addPage(pdf2pdf.getPage(page))
                        print('正在合并第{}页PDF文件'.format(pdf_page))
                        pdf_page += 1
                        pass
                    pass
                else:
                    pdf_finalWriter.addPage(pdf2pdf.getPage(0))
                    print('正在合并第{}页PDF文件'.format(pdf_page))
                    pdf_page += 1
                    pass
                pass
            with open(path_root + '\\' + 'output.pdf', 'wb') as f:
                print('正在整合所有PDF')
                pdf_finalWriter.write(f)
                print('已合并所有文件至output.pdf中')
                pass
        else:
            print('文件夹下没有文件')
            pass
        pass

    win32api.MessageBox(0, 'completed!', '提醒', win32con.MB_OK)
    pass


main()
