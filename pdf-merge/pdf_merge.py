#!/usr/bin/python
# -*- coding: UTF-8 -*-
# 利用PyPDF2模块合并同一文件夹下的所有PDF文件
# 只需修改存放PDF文件的文件夹变量：file_dir 和 输出文件名变量: outfile

import os
import fitz
from PyPDF2 import PdfFileReader, PdfFileWriter
from PIL import Image


# 使用os模块的walk函数，搜索出指定目录下的全部PDF文件
# 获取同一目录下的所有PDF文件的绝对路径
def getFileName(filedir):
    file_list = [os.path.join(root, filespath) \
                 for root, dirs, files in os.walk(filedir) \
                 for filespath in files \
                 if str(filespath).endswith('pdf')
                 ]
    return file_list if file_list else []


def MergeTwoPDF(first, source, outfile):
    output = PdfFileWriter()

    ## 读取图片pdf文件
    fInput = PdfFileReader(open(first, "rb"))
    ## 将第一页加入到输出文件
    output.addPage(fInput.getPage(0))

    ## 读原始文件
    input = PdfFileReader(open(source, "rb"))
    for iPage in range(1, input.getNumPages()):
        output.addPage(input.getPage(iPage))
    outputStream = open(outfile, "wb")
    output.write(outputStream)
    outputStream.close()
    print("PDF文件合并完成！")


# 合并同一目录下的所有PDF文件
def MergePDF(first_dir, source_dir):
    source_files = os.listdir(source_dir)
    if source_files:
        for pdf_file in source_files:
            print("路径：%s" % pdf_file)
            first = os.path.join(first_dir, pdf_file)
            source = os.path.join(source_dir, pdf_file)
            outfile = os.path.join("./result", pdf_file)
            MergeTwoPDF(first, source, outfile)
    else:
        print("没有可以修改的PDF文件！")


def TransPDFToImg(dir_path):
    files = os.listdir(dir_path)
    if files:
        for pdf_file in files:
            print("路径：%s" % pdf_file)
            file = os.path.join(dir_path, pdf_file)
            pdf_name = os.path.splitext(pdf_file)[0]
            image_path = os.path.join("./result", pdf_name)
            if not os.path.exists(image_path):
                os.mkdir(image_path)
            pdfImage(file, os.path.join(image_path, pdf_name + "_"), 5, 5, 0)
    else:
        print("没有可以转换的PDF文件！")


def pdfImage(pdf_path, img_path, zoom_x, zoom_y, rotation_angle):
    # 打开PDF文件
    pdf = fitz.open(pdf_path)
    # 逐页读取PDF
    for pg in range(0, pdf.pageCount):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotation_angle)
        pm = page.getPixmap(matrix=trans, alpha=False)
        # 开始写图像
        png_path = img_path + str(pg) + ".png"
        jpg_path = img_path + str(pg) + ".jpg"
        pm.writePNG(img_path + str(pg) + ".png")
        # png转jpg
        im = Image.open(png_path)
        im.save(jpg_path)
        os.remove(png_path)
    pdf.close()


def IsValidImage(img_path):
    """
    判断文件是否为有效（完整）的图片
    :param img_path:图片路径
    :return:True：有效 False：无效
    """
    bValid = True
    try:
        Image.open(img_path).verify()
    except:
        bValid = False
    return bValid


def transimg(img_path):
    """
    转换图片格式
    :param img_path:图片路径
    :return: True：成功 False：失败
    """
    if IsValidImage(img_path):
        try:
            str = img_path.rsplit(".", 1)
            output_img_path = str[0] + ".jpg"
            print(output_img_path)

            return True
        except:
            return False
    else:
        return False


if __name__ == '__main__':
    if not os.path.exists("./result"):
        os.mkdir("./result")
    while 1:
        fType = input("请输入序号选择需要的功能(exit or out 为退出):\n1:pdf合并\n2:pdf转为jpg\n")
        if fType == '1':
            source = input("请输入原pdf目录:")
            img = input("请输入图片pdf目录:")
            MergePDF(img, source)
        elif fType == '2':
            dir_path = input("请输入pdf目录:")
            TransPDFToImg(dir_path)
        elif fType == 'out' or fType == 'exit':
            break
        else:
            continue
