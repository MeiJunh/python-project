# -*- coding:utf-8 -*-
import csv
import fitz
import os
import re
import cmd
from PIL import Image
from pyzbar import pyzbar

HEADERS = ["原文件名", "新文件名", "发票代码", "发票号码", "开票日期", "校验码", "服务名称", "税额", "金额", "合计"]


class Fapiao(cmd.Cmd):
    """ 发票信息提取工具 """
    intro = '欢迎使用发票提取工具，输入?(help)获取帮助消息和命令列表，CTRL+C退出程序。\n'
    prompt = '\n输入命令: '
    doc_header = "详细文档 (输入 help <命令>):"
    misc_header = "友情提示:"
    undoc_header = "没有帮助文档:"
    nohelp = "*** 没有命令(%s)的帮助信息 "

    def do_load(self, arg):
        """ 加载发票 例如：load /path/to/invoices """
        arg = arg.strip()
        if not os.path.isdir(arg):
            print('参数必须是目录!')
            return

        pdfs = []
        for root, _, files in os.walk(arg):
            for fn in files:
                if fn.lower().endswith('.pdf'):
                    fpth = os.path.join(root, fn)
                    print(f'发现pdf文件: {fpth}')
                    pdfs.append(fpth)

        if not pdfs:
            print('未发现任何PDF文件。')
            return

        res_path = os.path.join(arg, "result.csv")
        self._parse_pdfs(pdfs, res_path)
        print(f'\n处理完成，结果已保存至: {res_path}')

    def _parse_pdfs(self, pdfs, res_path):
        """ 批量解析PDF发票并输出CSV """
        pic_dir = os.path.join(os.path.dirname(res_path), "pic")
        os.makedirs(pic_dir, exist_ok=True)

        infos = []
        for fpth in pdfs:
            try:
                info = self._parse_single(fpth, pic_dir)
                infos.append(info)
            except Exception as e:
                print(f'解析失败 [{fpth}]: {e}')

        with open(res_path, "w", encoding="utf-8", newline='') as fp:
            writer = csv.writer(fp)
            writer.writerow(HEADERS)
            writer.writerows(infos)

        # 打印结果预览
        self._print_preview(infos)

    def _parse_single(self, fpth, pic_dir):
        """ 解析单个PDF发票 """
        file_name = os.path.basename(fpth)
        pic_path = os.path.join(pic_dir, f"{file_name}.png")

        # 用 PyMuPDF 同时提取图片和文本
        qr_info = self._extract_qr(fpth, pic_path)
        money_info = self._extract_text(fpth)

        new_file_name = f"服务名_{qr_info[2]}_{qr_info[0]}_{qr_info[1]}_{money_info[3]}.pdf"
        return [file_name, new_file_name] + qr_info + money_info

    def _extract_qr(self, pdf_path, pic_path):
        """ 从PDF中提取第一张图片并解码二维码 """
        doc = fitz.open(pdf_path)
        try:
            for page in doc:
                for img_info in page.get_images(full=True):
                    xref = img_info[0]
                    pix = fitz.Pixmap(doc, xref)
                    # CMYK 转 RGB
                    if pix.n >= 5:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    pix.save(pic_path)
                    # 解码二维码
                    result = self._decode_qr(pic_path)
                    if result:
                        return result
        finally:
            doc.close()
        return ['', '', '', '']

    def _decode_qr(self, img_path):
        """ 解码二维码图片 """
        if not os.path.isfile(img_path):
            return None
        img = Image.open(img_path)
        txt_list = pyzbar.decode(img)
        if not txt_list:
            return None
        parts = txt_list[0].data.decode("utf-8").split(",")
        if len(parts) < 7:
            return None
        # 发票代码, 发票号码, 开票日期, 校验码
        return [parts[2], parts[3], parts[5], parts[6]]

    def _extract_text(self, pdf_path):
        """ 用 PyMuPDF 提取文本，解析服务名称和金额信息 """
        doc = fitz.open(pdf_path)
        money = []
        service = ""

        try:
            for page in doc:
                page_width = page.rect.width
                blocks = page.get_text("blocks")
                for block in blocks:
                    # block: (x0, y0, x1, y1, text, block_no, block_type)
                    if block[6] != 0:  # 跳过非文本块
                        continue
                    text = block[4].strip()
                    x0 = block[0]

                    if "*" in text and x0 < page_width / 2:
                        text_clean = text.replace("\n", " ")
                        idx = text_clean.index("*")
                        service += text_clean[idx:]
                    elif "￥" in text or "¥" in text:
                        money.append(text.replace("\n", "").replace(" ", ""))
        finally:
            doc.close()

        money.sort()
        # 补齐为 [税额, 金额, 合计]
        while len(money) < 3:
            money.insert(0, "0")
        return [service] + money[:3]

    def _print_preview(self, infos):
        """ 在终端打印结果预览 """
        if not infos:
            return
        print(f'\n{"="*60}')
        print(f'共处理 {len(infos)} 张发票:')
        print(f'{"="*60}')
        for info in infos:
            print(f'  {info[0]} -> {info[1]}')
            print(f'    发票代码: {info[2]}  发票号码: {info[3]}  合计: {info[9]}')
        print(f'{"="*60}')


if __name__ == '__main__':
    try:
        Fapiao().cmdloop()
    except KeyboardInterrupt:
        print('\n\n再见！')
