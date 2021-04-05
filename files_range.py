#!/usr/bin/env python
# _*_ coding: utf-8 _*_
# @Time : 2021-04-01 10:41
# @Author : Pan
# @Version：V 0.1
# @File : files_range.py
# @desc :


import os
import re
import shutil
from pathlib import Path

from PyPDF2 import PdfFileReader


class ElecInvoice(object):

    def file_classify(self, path, file_path):
        """
        func:将文件按照后缀名分类存入对应文件夹下
        path:待处理文件所在的根目录
        """
        file = os.listdir(path)  # 列出当前文件夹的所有文件
        for f in file:  # 循环遍历每个文件
            folder_name = file_path + f.split(".")[-1]  # 以扩展名为名称的子文件夹
            if not os.path.exists(folder_name):  # 如果不存在该目录
                os.makedirs(folder_name)  # 先创建，再移动文件
                shutil.move(f, folder_name)
            else:  # 直接移动文件
                shutil.move(f, folder_name)

    def move_file(self, scr_path):
        """
        func:将文件按照文件类型存入对应文件夹下
        path:待处理文件所在的根目录
        """
        FILE_FORMATS = {
            "图片资料": [".jpeg", ".jpg", ".tiff", ".gif", ".bmp", ".png", ".bpg", "svg", ".heif", ".psd"],
            "文档资料": [".oxps", ".epub", ".pages", ".docx", ".doc", ".fdf", ".ods", ".odt", ".pwi", ".xsn",
                     ".xps", ".dotx", ".docm", ".dox", ".rvg", ".rtf", ".rtfd", ".wpd", ".xls", ".xlsx",
                     ".xlsm", ".ppt", ".pptx", ".csv", ".pdf", ".md", ".xmind"],
            "视频文件": [".avi", ".flv", ".wmv", ".mov", ".mp4", ".webm", ".vob", ".mng", ".qt", ".mpg", ".mpeg", ".3gp",
                     ".mkv"],
            "压缩文件": [".a", ".ar", ".cpio", ".iso", ".tar", ".gz", ".rz", ".7z", ".dmg", ".rar", ".xar", ".zip"],
            "程序文件": [".exe", ".bat", ".lnk", ".js", ".dll", ".db", ".py", ".html5", ".html", ".htm", ".xhtml",
                     ".cpp", ".java", ".css", ".sql", ".msi"],
            "网页文件": ['.html', '.xml', '.mhtml', '.html'],
            "音频文件": [".aac", ".aa", ".aac", ".dvf", ".m4a", ".m4b", ".m4p", ".mp3", ".msv", ".ogg", ".oga",
                     ".raw", ".vox", ".wav", ".wma"],
        }
        for my_file in Path(scr_path).iterdir():
            # is_dir()判定是否为目录
            if my_file.is_dir():
                # 用continue就跳过了文件夹
                continue
            else:
                file_path = Path(scr_path + '\\' + my_file.name)  # 拼接形成文件
                lower_file_path = file_path.suffix.lower()  # 后缀转化成小写
                for my_key in FILE_FORMATS:
                    if lower_file_path in FILE_FORMATS[my_key]:  # 如果后缀名在上面定义的
                        directory_path = Path(scr_path + '\\' + my_key)
                        print(directory_path)
                        # 如果文件夹不存在，则根据定义建立文件夹
                        directory_path.mkdir(exist_ok=True)
                        file_path.rename(directory_path.joinpath(my_file.name))
        print('文件分类已结束！')

    # 递归遍历文件，将文件名写入空list[]all_file
    def traverse_file(self, scr_path, all_files):
        # 首先遍历当前目录所有文件及文件夹
        file_list = os.listdir(scr_path)
        # 准备循环判断每个元素是否是文件夹还是文件，是文件的话，把名称传入list，是文件夹的话，递归
        for file_name in file_list:
            cur_path = os.path.join(scr_path, file_name)
            if os.path.isdir(cur_path):
                self.traverse_file(cur_path, all_files)
            else:
                all_files.append(cur_path)
        return all_files

    # 遍历文件名，筛选出电子发票移动到目标目录
    def show_pdf_files(self, all_files_list, dst_path):
        """
        all_files_list:所有文件的列表
        dst_path:电子发票移动目录
        """
        for file_name in all_files_list:
            try:
                if file_name.endswith(".pdf"):  # 判断是否是PDF文件
                    height, width = self.run_pdf_size(file_name)
                    list1 = file_name.split('\\')
                    new_file_name = os.path.join(dst_path, list1[-1])
                    # 通过识别PDF第一页尺寸判断文件是否是电子发票
                    if 390 < height < 400 and 590 < width < 650:
                        print(f"这是一个电子发票：{file_name}，页面尺寸：{width, height}，移动到：{dst_path}")
                        shutil.move(file_name, new_file_name)  # 移动文件
                    else:
                        print(f"提示：非电子发票：{file_name}，页面尺寸：{width, height}")
                else:
                    continue
            except Exception as e:
                print(e)
        return

    def run_pdf_size(self, filename):
        """
        func:获取PDF页面尺寸
        filename:pdf文件的名称
        """
        try:
            with open(filename, 'rb') as f:
                pdf = PdfFileReader(f)
                page_1 = pdf.getPage(0)
                if page_1.get('/Rotate', 0) in [90, 270]:
                    return page_1['/MediaBox'][2], page_1['/MediaBox'][3]
                else:
                    return page_1['/MediaBox'][3], page_1['/MediaBox'][2]
        except Exception as e:
            print(e)

    def auto_save_file(self, file_name):
        """
        func:判断文件重名
        file_name:文件名称
        """
        directory, file_name = os.path.split(file_name)
        while os.path.isfile(file_name):
            pattern = '(\d+)\)\.'
            if re.search(pattern, file_name) is None:
                file_name = file_name.replace('.', '(0).')
            else:
                current_number = int(re.findall(pattern, file_name)[-1])
                new_number = current_number + 1
                file_name = file_name.replace(f'({current_number}).', f'({new_number}).')
            file_name = os.path.join(directory + os.sep + file_name)
        return file_name


if __name__ == '__main__':
    ei = ElecInvoice()
