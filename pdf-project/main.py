# -*- coding:utf-8 -*-
import csv

import fitz
import os.path, re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.pdfpage import PDFTextExtractionNotAllowed, PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from PIL import Image
from pyzbar import pyzbar
import cmd

head = ["文件名", "发票代码", "发票号码", "开票日期", "校验码", "服务名称", "税额", "金额", "合计"]


class FapiaoShell(cmd.Cmd):
    """ 发票 """
    intro = '欢迎使用发票提取工具，输入?(help)获取帮助消息和命令列表，CTRL+C退出程序。\n'
    prompt = '\n输入命令: '
    doc_header = "详细文档 (输入 help <命令>):"
    misc_header = "友情提示:"
    undoc_header = "没有帮助文档:"
    nohelp = "*** 没有命令(%s)的帮助信息 "

    def __init__(self):
        super().__init__()

    def do_load(self, arg):
        """ 加载发票 例如：load D:\ """
        if not os.path.isdir(arg):
            print('参数必须是目录!')
            return
        os.chdir(os.path.dirname(arg))
        pdfs = []
        for root, _, files in os.walk(arg):
            for fn in files:
                ext = os.path.splitext(fn)[1].lower()
                if ext != '.pdf':
                    continue
                fpth = os.path.join(root, fn)
                fpth = os.path.relpath(fpth)
                print(f'发现pdf文件: {fpth}')
                pdfs.append(fpth)
        self._parse_pdfs(pdfs, os.path.join(arg, "result.csv"))

    def _parse_pdfs(self, pdfs, res_path):
        """ 分析 """
        # 结果使用csv文件保存
        fp = open(res_path, "w", encoding="utf-8")
        writer = csv.writer(fp)
        writer.writerow(head)
        if not os.path.exists("./pic"):
            os.makedirs("./pic")
        for fpth in pdfs:
            tmp_info = []
            # 获取当前文件名
            file_name = fpth.split("/")[-1]
            # 该pdf对应的二维码存储的文件名
            pic_name = os.path.join("./pic", "{}.png".format(file_name))

            # 将二维码从pdf中分离
            self._pdf2pic(fpth, pic_name)
            # 获取二维码信息得到的信息，分别为"发票代码", "发票号码", "开票日期", "校验码"
            pic_info = self._get_ewm(pic_name)
            # 通过解析pdf，正则表达解析的方式获取"服务名称", 税额","金额", "合计"
            money_info = self._extrace_from_words(fpth)
            # 将信息按照head一次填入tmp_info中
            tmp_info.append(file_name)
            tmp_info.extend(pic_info)
            tmp_info.extend(money_info)

            writer.writerow(tmp_info)
            # 将每张图片结果都写入文件中
        fp.close()

    def _extrace_from_words(self, pdf_path):
        # 以二进制读模式打开
        file = open(pdf_path, 'rb')
        # 用文件对象来创建一个pdf文档分析器
        parser = PDFParser(file)
        # 创建一个PDF文档对象存储文档结构,提供密码初始化,没有就不用传该参数
        doc = PDFDocument(parser, password='')
        # 检查文件是否允许文本提取
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        # 创建PDf资源管理器来管理共享资源，#caching = False不缓存
        rsrcmgr = PDFResourceManager(caching=False)
        # 创建一个PDF设备对象
        laparams = LAParams()
        # 创建一个PDF页面聚合对象
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF解析器对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        # 获得文档的目录(纲要),文档没有纲要会报错
        # PDF文档没有目录时会报：raise PDFNoOutlines  pdfminer.pdfdocument.PDFNoOutlines
        # print(doc.get_outlines())

        # 获取page列表
        # print(PDFPage.get_pages(doc))
        # 循环遍历列表，每次处理一个page的内容
        # money存储税额、金额、合计
        money = []
        # data存储服务名称
        data = ""
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)
            # 接受该页面的LTPage对象
            layout = device.get_result()
            # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
            # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal等等
            for x in layout:
                # # 如果x是水平文本对象的话
                if (isinstance(x, LTTextBoxHorizontal)):
                    # print('kkk')
                    text_info = x.get_text()
                    # 获取的
                    if "*" in text_info:
                        if x.x0 > (layout.x1 / 2):
                            continue
                        service_info = text_info.replace("\n", " ")
                        index = service_info.index("*")
                        data += service_info[index:]

                    elif "￥" in text_info or "¥" in text_info:
                        money.append(text_info.replace("\n", "").replace(" ", ""))
        info = []
        info.append(data)
        money.sort()
        money_len = len(money)
        money_info = []
        if money_len >= 3:
            money_info = money[:3]
        elif money_len == 2:
            money_info = [0, money[0], money[1]]
        elif money_len == 1:
            money_info = [0, 0, money[0]]
        elif money_len == 0:
            money_info = [0, 0, 0]
        info.extend(money_info)
        return info

    # 获取图片
    def _pdf2pic(self, path, pic_path):
        checkXO = r"/Type(?= */XObject)"  # 使用正则表达式来查找图片
        checkIM = r"/Subtype(?= */Image)"
        doc = fitz.open(path)  # 打开pdf文件
        imgcount = 0  # 图片计数
        lenXREF = doc._getXrefLength()  # 获取对象数量长度

        # 遍历每一个对象
        for i in range(1, lenXREF):
            text = doc._getXrefString(i)  # 定义对象字符串
            isXObject = re.search(checkXO, text)  # 使用正则表达式查看是否是对象
            isImage = re.search(checkIM, text)  # 使用正则表达式查看是否是图片
            if not isXObject or not isImage:  # 如果不是对象也不是图片，则continue
                continue
            imgcount += 1
            pix = fitz.Pixmap(doc, i)  # 生成图像对象
            if pix.n < 5:  # 如果pix.n<5,可以直接存为PNG
                pix.writePNG(pic_path)
            else:  # 否则先转换CMYK
                pix0 = fitz.Pixmap(fitz.csRGB, pix)
                pix0.writePNG(pic_path)
                pix0 = None
            pix = None  # 释放资源
            break

    def _get_ewm(self, img_adds):
        """ 读取二维码的内容： img_adds：二维码地址（可以是网址也可是本地地址 """
        if os.path.isfile(img_adds):
            # 从本地加载二维码图片
            img = Image.open(img_adds)

        # 获取到二维码对应的信息
        txt_list = pyzbar.decode(img)
        if len(txt_list) <= 0:
            return ['', '', '', '']
        decode_info = txt_list[0].data.decode("utf-8").split(",")
        info = []
        info.append(decode_info[2])
        info.append(decode_info[3])
        info.append(decode_info[5])
        info.append(decode_info[6])
        return info

if __name__ == '__main__':
    try:
        FapiaoShell().cmdloop()
    except KeyboardInterrupt:
        print('\n\n再见！')
