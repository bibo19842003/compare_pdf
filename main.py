import os
import sys
import shutil
import datetime
import threading
from itertools import product
import math
from typing import Union, Tuple
import json

import fitz
from paddleocr import PaddleOCR
from difflib import HtmlDiff
import customtkinter
from CTkMessagebox import CTkMessagebox
from win32com import client
import pymupdf
from pikepdf import Pdf,Page,Rectangle
from reportlab.lib import units
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import  TTFont

from config import pdf_font_dict





class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        customtkinter.deactivate_automatic_dpi_awareness()
        self.textbox_font = customtkinter.CTkFont(family="微软雅黑", size=12)
        self.entry_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.button_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.msg_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.radio_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.combox_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.font = customtkinter.CTkFont(family="微软雅黑", size=22)
        self.label_font = customtkinter.CTkFont(family="Arial", size=24)
        self.title("PDF文本比较工具 v1.3.0_beta")
        self.geometry(("750x500"))
        self.resizable(0,0)
        self.configure(fg_color="#FDE6E0")
        
        self.file_directory = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        
        # msgbox configure
        # bg_color=self.msg_bg_color, fg_color=self.msg_fg_color, 
        self.msg_bg_color = "#95e1d3"
        self.msg_fg_color = "#f38181"
        self.msg_width = 300
        self.msg_height = 150

        # paddle ocr 初始化
        det_model_dir = os.path.join(self.file_directory, "ch_PP-OCRv4_det_infer")
        rec_model_dir = os.path.join(self.file_directory, "ch_PP-OCRv4_rec_infer")
        cls_model_dir = os.path.join(self.file_directory, "ch_ppocr_mobile_v2.0_cls_infer")
        self.ocr = PaddleOCR(use_angle_cls=True, lang="ch", use_gpu=False, det_model_dir=det_model_dir, rec_model_dir=rec_model_dir, cls_model_dir=cls_model_dir)
        
        # pdf 字体类型加载
        self.pdf_watermark_template = os.path.join(self.file_directory, "pdf_watermark_template.pdf")
        self.sys_font_list = [ i for i in os.listdir(r"C:\Windows\Fonts")]
        self.pdf_font_list = [ k for k, v in pdf_font_dict.items() if v in self.sys_font_list]
        for k in self.pdf_font_list:
            pdfmetrics.registerFont(TTFont(k, os.path.join("C:\Windows\Fonts", pdf_font_dict[k])))

        
        self.user_ui()
        

    # 范围：pdf文本比较
    # 选择文件夹
    def select_folder(self, index):
        folder = customtkinter.filedialog.askopenfilename()
        if folder:
            if index == 1:
                self.fc_entry_select_file_1.delete(0, customtkinter.END)
                self.fc_entry_select_file_1.insert(0, folder)
            elif index == 2:
                self.fc_entry_select_file_2.delete(0, customtkinter.END)
                self.fc_entry_select_file_2.insert(0, folder)
            

    # 范围：pdf文本比较
    # 获取一个文件夹内的所有图片的文本
    def paddleocr_get_mul_pic_text(self, ocr, img_folder, file):
        img_list = []
        for i in os.listdir(img_folder):
            if i.endswith(".png"):
                img_list.append(i)
            
        img_sort_list = sorted(img_list)
    
        total_res = ""
        for index, j in enumerate(img_sort_list):
            image_file = os.path.join(self.file_directory, img_folder, j)
            self.status_message_add(f"        {file} 第 {index+1} 页去水印 启动")
            result = ocr.ocr(image_file, cls=True)
            total_res += f"以下是第 {index+1} 页内容 \n"
            for idx in range(len(result)):
                res = result[idx]
                for line in res:
                    res_line = line[1][0] + "\n"
                    total_res += res_line
            total_res += "\n"
            
        with open(os.path.join(img_folder, "text_content.txt"), "w", encoding="utf-8") as f:
            f.write(total_res)
        
        
    # 范围：pdf文本比较
    # 获取单个图片内的文本
    def paddleocr_get_single_pic_text(self, ocr, img_folder, file, index):
        image_file = os.path.join(self.file_directory, img_folder, "pdf_split_" + str(index+1).zfill(3) + ".png")
        self.status_message_add(f"        {file} 第 {index+1} 页获取去水印后的文件内容")
        result = ocr.ocr(image_file, cls=True)
        page_res = f"以下是第 {index+1} 页内容 \n"
        for idx in range(len(result)):
            res = result[idx]
            for line in res:
                res_line = line[1][0] + "\n"
                page_res += res_line
        page_res += "\n"
            
        return page_res


    # 范围：pdf文本比较
    # 单张图片去水印
    def mt_pic_remove_watermark(self, index, page, pdf_file, img_folder, print_compare_status_msg=True, wm_threshold=600):
            rotate = int(0)
            # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高4倍的图像
            # zoom_x, zoom_y = 2, 2
            zoom_x, zoom_y = 2, 2
            trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
            pixmap = page.get_pixmap(matrix=trans, alpha=False)
            for pos in product(range(pixmap.width), range(pixmap.height)):
                rgb = pixmap.pixel(pos[0], pos[1])
                if (sum(rgb) >= wm_threshold):
                    pixmap.set_pixel(pos[0], pos[1], (255, 255, 255))
            pixmap.pil_save(os.path.join(img_folder, "pdf_split_" + str(index+1).zfill(3) + ".png"))
            print(f"    {pdf_file} 第 {index+1} 页水印去除完成")
            if print_compare_status_msg:
                self.status_message_add(f"        {pdf_file} 第 {index+1} 页去水印完成")
            

    # 范围：pdf文本比较
    # pdf 每一页转图片, 然后图片去水印
    def pdf_to_pic_remove_watermark(self, pdf_file, img_folder):
        t_list = []
        pdf = fitz.open(pdf_file)
        
        for index, page in enumerate(pdf):
            self.mt_pic_remove_watermark(index, page, pdf_file, img_folder)
        

    # 范围：pdf文本比较
    # 获取 文本文件内容
    def get_file_content(self, file_path):
        lines = []
        with open(file_path, mode="r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        return lines 
    

    # 范围：pdf文本比较
    # 文本内容进行比较，生成diff文件
    def compare_file(self, file1, file2):
        lines1 = self.get_file_content(file1)
        lines2 = self.get_file_content(file2)

        # 找出差异输出到result(str)
        html_diff = HtmlDiff()
        result = html_diff.make_file(lines1, lines2)
    
        # 将差异写入html文件
        with open("comparison.html", "w", encoding="utf-8") as f:
            f.write(result)
            

    # 范围：pdf文本比较
    # auto 模式获取 pdf 文件的内容
    def get_pdf_auto(self, pdf_file, img_folder):
        total_res = ""
        open_file = fitz.open(pdf_file)
        
        for index, page in enumerate(open_file):
            self.status_message_add(f"    {pdf_file} 第 {index+1} 页 内容获取开始")
            # 判断是否有图片
            list_image = page.get_images()
            if list_image:
                # 去水印
                self.mt_pic_remove_watermark(index, page, pdf_file, img_folder)
                # 获取内容
                img_text = self.paddleocr_get_single_pic_text(self.ocr, img_folder, pdf_file, index)
                total_res = total_res + img_text + "\n"
            else:
                #print("No images found on page", index)
                total_res = total_res + f"以下是第 {index+1} 页内容 \n" + page.get_text() + "\n"
                
            self.status_message_add(f"    {pdf_file} 第 {index+1} 页 内容获取结束")
            
        with open(os.path.join(img_folder, "text_content.txt"), "w", encoding="utf-8") as f:
            f.write(total_res)
            

    # 范围：pdf文本比较
    # 文本模式获取 pdf 文件的内容
    def get_pdf_text(self, pdf_file, img_folder):
        total_res = ""
        open_file = fitz.open(pdf_file)
        
        for index, page in enumerate(open_file):
            self.status_message_add(f"    {pdf_file} 第 {index+1} 页 内容获取开始")
            total_res = total_res + f"以下是第 {index+1} 页内容 \n" + page.get_text() + "\n"
                
            self.status_message_add(f"    {pdf_file} 第 {index+1} 页 内容获取结束")
            
        with open(os.path.join(img_folder, "text_content.txt"), "w", encoding="utf-8") as f:
            f.write(total_res)
    

    # 范围：pdf文本比较
    # 进行相关检验，然后对文本内容进行比较，生成diff文件
    def compare_and_create(self, file1, file2):
        try:
        #if 1:
            self.fc_button_compare_file.configure(state="disabled")

            self.status_message_init("---运行开始---")
        
            # 1 检查是否选择文件
            if not os.path.isfile(file1) or not os.path.isfile(file1):
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='请选择要比较的文件！')
                return
        
            # 2 图片文件夹初始化
            temp_img_folder_1 = os.path.join(self.file_directory, "temp_img_folder_1")
            temp_img_folder_2 = os.path.join(self.file_directory, "temp_img_folder_2")

            if os.path.isdir(temp_img_folder_1):
                shutil.rmtree(temp_img_folder_1)
            os.mkdir(temp_img_folder_1)

            if os.path.isdir(temp_img_folder_2):
                shutil.rmtree(temp_img_folder_2)
            os.mkdir(temp_img_folder_2)

            # 3 运行模式判断，获取 pdf 文件的内容
            if self.fc_mode_var.get() == "mode_auto":
                self.status_message_add(f"获取 {file1} 文件内容开始")
                self.get_pdf_auto(file1, temp_img_folder_1)
                self.status_message_add(f"获取 {file1} 文件内容结束")
                self.status_message_add(f"获取 {file2} 文件内容开始")
                self.get_pdf_auto(file2, temp_img_folder_2)
                self.status_message_add(f"获取 {file2} 文件内容结束")
            elif self.fc_mode_var.get() == "mode_text":
                self.status_message_add(f"获取 {file1} 文件内容开始")
                self.get_pdf_text(file1, temp_img_folder_1)
                self.status_message_add(f"获取 {file1} 文件内容结束")
                self.status_message_add(f"获取 {file2} 文件内容开始")
                self.get_pdf_text(file2, temp_img_folder_2)
                self.status_message_add(f"获取 {file2} 文件内容结束")
            elif self.fc_mode_var.get() == "mode_image":
                # pdf 去水印
                self.status_message_add(f"启动 去水印 操作")
                self.pdf_to_pic_remove_watermark(file1, temp_img_folder_1)
                self.pdf_to_pic_remove_watermark(file2, temp_img_folder_2)
        
                # 使用paddle ocr获取去水印后的文本
                self.status_message_add(f"启动 获取去水印后的文件内容 操作")
                self.paddleocr_get_mul_pic_text(self.ocr, temp_img_folder_1, file1)
                self.paddleocr_get_mul_pic_text(self.ocr, temp_img_folder_2, file2)


            # 4 文件比较
            self.compare_file(os.path.join(temp_img_folder_1, "text_content.txt"), os.path.join(temp_img_folder_2, "text_content.txt"))
            print("\n'comparison.html' 已在程序运行的目录生成，请用浏览器打开。")
            self.status_message_add("\n'comparison.html' 已生成，请用浏览器打开。生成文件在程序运行的目录下查看。")
            CTkMessagebox(title='成功', font=self.msg_font, justify="center", option_1="确定", icon="check", width=self.msg_width, height=self.msg_height, message="'comparison.html' 已在程序运行的目录生成，请用浏览器打开。")
            self.fc_button_compare_file.configure(state="normal")
        except:
            print("程序运行异常，请反馈给工具维护人员。")
            self.status_message_add("程序运行异常，请反馈给工具维护人员。\n程序运行异常，请反馈给工具维护人员。\n程序运行异常，请反馈给工具维护人员。\n")
            CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message="程序运行异常，请反馈给工具维护人员。")
            self.fc_button_compare_file.configure(state="normal")


    # 范围：pdf文本比较
    # 状态栏打印信息 初始化
    def status_message_init(self, message):
        self.fc_textbox_log.configure(state=customtkinter.NORMAL)
        self.fc_textbox_log.delete(1.0, customtkinter.END)
        self.fc_textbox_log.insert(customtkinter.END, message + "\n")
        self.fc_textbox_log.configure(state=customtkinter.DISABLED)
        

    # 范围：pdf文本比较
    # 状态栏打印信息 增加
    def status_message_add(self, message):
        self.fc_textbox_log.configure(state=customtkinter.NORMAL)
        self.fc_textbox_log.insert(customtkinter.END, message + "\n")
        self.fc_textbox_log.configure(state=customtkinter.DISABLED)
        self.fc_textbox_log.see(customtkinter.END)
        

    # 范围：文件格式转换
    # 选择文件
    def select_convert_file(self, index):
        folder = customtkinter.filedialog.askopenfilename()
        if folder:
            self.ft_entry_select_file.delete(0, customtkinter.END)
            self.ft_entry_select_file.insert(0, folder)


    # 范围：文件格式转换
    # 文件格式转换
    def word2pdf(self, file):
        try:
            # 2 后缀名不正确，1 成功，0 异常
            filepath, filename = os.path.split(file)
            suffix = filename.split(".")[-1]
            if suffix.lower() not in ["doc", "docx"]:
                return 2
            
            time_now = datetime.datetime.now()
            time_str = time_now.strftime('%Y%m%d_%H%M%S')
            new_file = os.path.join(filepath, filename.split(".")[0] + time_str + ".pdf")
            
            word = client.Dispatch("Word.Application")
            doc = word.Documents.Open(file)  # 打开word文件
            doc.SaveAs(new_file, 17)
            doc.Close()
            word.Quit()
            return 1
        except:
            return 0
    
    
    # 范围：文件格式转换
    # 文件格式转换
    def file_convert(self, file):
        try:
            self.ft_button_convert.configure(state="disabled")

            # 1 检查是否选择文件
            if not os.path.isfile(file):
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='请选择要转换的文件！')
                return
        
            # 2 文件格式转换
            if self.ft_mode_var.get() == "word2pdf":
                r = self.word2pdf(file)
                if r == 2:
                    CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='文件后缀名不正确，需为 doc 或 docx')
                elif r == 0:
                    CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='文件格式转换失败')
                elif r == 1:
                    CTkMessagebox(title='成功', font=self.msg_font, justify="center", option_1="确定", icon="check", width=self.msg_width, height=self.msg_height, message="转换成功，请在程序运行目录查看生成文件。")
            self.ft_button_convert.configure(state="normal")
        except:
            CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message="程序运行异常，请反馈给工具维护人员。")
            self.ft_button_convert.configure(state="normal")


    # 范围：水印操作
    # 选择文件
    def wh_select_convert_file(self, index):
        file_types = [("pdf Files", "*.pdf"), ("PDF Files", "*.PDF")]
        folder = customtkinter.filedialog.askopenfilename(filetypes=file_types, title="请选择pdf文件")
        if folder:
            self.wh_entry_select_file.delete(0, customtkinter.END)
            self.wh_entry_select_file.insert(0, folder)
            
            
    # 范围：水印操作        
    # 创建 字符串 水印 pdf文件
    def create_wartmark(self,
                        content:str,
                        filename:str,
                        width: Union[int, float],
                        height: Union[int, float],
                        font: str,
                        fontsize: int,
                        x_offset: int,
                        y_offset: int,
                        angle: Union[int, float] = 45,
                        text_stroke_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                        text_fill_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                        text_fill_alpha: Union[int, float] = 1) -> None:

        #创建PDF文件，指定文件名及尺寸，以像素为单位
        c = canvas.Canvas(filename, pagesize=(width, height))

        #画布平移保证文字完整性
        str_len = fontsize * len(content)*3/5
        c.translate(int((width - str_len*math.cos(math.radians(angle)))/2)+x_offset, int(height/2) - int(str_len * math.sin(math.radians(angle))/2)+y_offset)

        #设置旋转角度
        c.rotate(angle)

        #设置字体大小
        c.setFont(font,fontsize)

        #设置字体轮廓彩色
        c.setStrokeColorRGB(*text_stroke_color_rgb)

        #设置填充色
        c.setFillColorRGB(*text_fill_color_rgb)

        #设置字体透明度
        c.setFillAlpha(text_fill_alpha)

        #绘制字体内容
        c.drawString(0,0,content)

        #保存文件
        c.save()
        

    # 范围：水印操作
    # 生成目标pdf水印文件
    def add_watemark(self,
                     target_pdf_path:str,
                     watermark_pdf_path:str,
                     nrow:int,
                     ncol:int) -> None:

        #选择需要添加水印的pdf文件
        target_pdf = Pdf.open(target_pdf_path)

        #读取水印pdf文件并提取水印
        watermark_pdf = Pdf.open(watermark_pdf_path)
        watermark_page = watermark_pdf.pages[0]

        #遍历目标pdf文件中的所有页，批量添加水印
        for idx,target_page in enumerate(target_pdf.pages):
            for x in range(ncol):
                for y in range(nrow):
                    #向目标页指定范围添加水印
                    target_page.add_overlay(watermark_page,
                                            Rectangle(target_page.trimbox[2] * x / ncol,
                                            target_page.trimbox[3] * y / nrow,
                                            target_page.trimbox[2] * (x + 1) / ncol,
                                            target_page.trimbox[3] * (y + 1) / nrow
                                            ))
        #保存PDF文件
        target_pdf.save(target_pdf_path[:-4] + '_已添加水印.pdf')
        

    # 范围：水印操作
    # 创建水印预览
    def pdf_create_del_watermark(self):
        select_pdf = self.wh_entry_select_file.get()
        if not os.path.isfile(select_pdf):
            CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='请选择要处理的PDF文件！')
            return

        # 判断添加或者删除水印
        if self.wh_mode_add_del_var.get() == "add":
            # 添加水印
            # 获取页面各种变量
            watermark_content = self.wh_entry_w_c.get()
            if watermark_content == "":
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='请填写水印内容！')
                return
                
            select_font_type = self.wh_combobox_font_type.get()
            
            try:
                select_font_size = int(self.wh_combobox_font_size.get())
            except:
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='字体大小选择有误，请重新选择！')
                return

            wm_row_num = int(self.wh_combobox_w_row.get())
            
            wm_col_num = int(self.wh_combobox_w_col.get())
            
            try:
                wm_angle = int(self.wh_combobox_w_angle.get())
            except:
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='水印角度选择有误，请重新选择！')
                return
            
            try:
                wm_offset_x = int(self.wh_combobox_w_offset_x.get())
            except:
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='水印偏移x 选择有误，请重新选择！')
                return
                
            try:
                wm_offset_y = int(self.wh_combobox_w_offset_y.get())
            except:
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='水印偏移y 选择有误，请重新选择！')
                return

            wm_transparency = float(self.wh_combobox_w_transparency.get())

            self.wh_button_pdf.configure(state="disabled")
            # 获取所选择的pdf文件的尺寸
            try:
                with pymupdf.open(select_pdf) as doc:
                    page = doc[0]
                    page_width = page.rect.width
                    page_height = page.rect.height
            except:
                CTkMessagebox(title='错误', font=self.msg_font, justify="center", option_1="退出", icon="cancel", width=self.msg_width, height=self.msg_height, message='所选择的pdf文件有问题，请重新选择！')
                self.wh_button_pdf.configure(state="normal")
                return

            # 创建 pdf 水印模板
            self.create_wartmark(content=watermark_content,
                                 filename=self.pdf_watermark_template,
                                 width=page_width,
                                 height=page_height,
                                 font=select_font_type,
                                 fontsize=select_font_size,
                                 x_offset=wm_offset_x,
                                 y_offset=wm_offset_y,
                                 angle=wm_angle,
                                 text_fill_alpha=wm_transparency)
            
            # 添加水印        
            self.add_watemark(target_pdf_path=select_pdf,
                              watermark_pdf_path=self.pdf_watermark_template,
                              nrow = wm_row_num,
                              ncol = wm_col_num)
                                 
            CTkMessagebox(title='水印添加成功', font=self.msg_font, justify="center", option_1="确定", icon="check", width=self.msg_width, height=self.msg_height, message="水印添加成功。请在所选择的pdf文件目录查看！")
            self.wh_button_pdf.configure(state="normal")
        else:
            # 删除水印
            # 判断图片模式 还是 纯文本模式
            self.wh_button_pdf.configure(state="disabled")
            
            wm_threshold = int(self.wh_combobox_w_threshold.get())

            # 图片模式
            temp_img_folder_remove_wm_1 = os.path.join(self.file_directory, "temp_img_folder_remove_wm_1")
            if os.path.isdir(temp_img_folder_remove_wm_1):
                shutil.rmtree(temp_img_folder_remove_wm_1)
            os.mkdir(temp_img_folder_remove_wm_1)

            open_file = fitz.open(select_pdf)
            for index, page in enumerate(open_file):
                # 去水印
                self.mt_pic_remove_watermark(index, page, select_pdf, temp_img_folder_remove_wm_1, print_compare_status_msg=False, wm_threshold=wm_threshold)
            self.pic_2_pdf_for_dir(temp_img_folder_remove_wm_1, select_pdf)
            CTkMessagebox(title='水印删除成功', font=self.msg_font, justify="center", option_1="确定", icon="check", width=self.msg_width, height=self.msg_height, message="水印删除成功。请在所选择的pdf文件目录查看！")

            self.wh_button_pdf.configure(state="normal")


    # 范围：水印操作
    # 文件夹下的图片转成一个pdf
    def pic_2_pdf_for_dir(self, img_folder, select_pdf):
        pdf = fitz.open()
        img_files = sorted(os.listdir(img_folder), key=lambda x: str(x).split('.')[0])
        img_type = ["png", "jpg", "jpeg", "bmp"]
        for img in img_files:
            if img.split(".")[-1].lower() in img_type:
                imgdoc = fitz.open(os.path.join(img_folder, img))
                #将打开后的图片转成单页pdf
                pdfbytes = imgdoc.convert_to_pdf()
                imgpdf = fitz.open("pdf", pdfbytes)
                #将单页pdf插入到新的pdf文档中
                pdf.insert_pdf(imgpdf)
        pdf.save(select_pdf[:-4] + '_已移除水印.pdf')
        pdf.close()


    # 范围：窗口
    def on_closing(self):
            self.destroy()
            self.quit()
            sys.exit()


    # 范围：窗口
    # UI 页面
    def user_ui(self):
        # 标签页
        customtkinter.CTkTabview._outer_button_overhang = 15
        self.tabview01 = customtkinter.CTkTabview(master=self, width=755, height=505, fg_color="#FDE6E0", segmented_button_fg_color="#FDE6E0", segmented_button_selected_color="#4F86EC", segmented_button_unselected_color="gray", border_width=2, corner_radius=0, bg_color="#FDE6E0", segmented_button_font=self.button_font)
        self.tabview01.place(x=-2, y=5)

        # 添加2个选项卡
        tabview01_title1 = "PDF文本比较"
        tabview01_title2 = "文本转换"
        tabview01_title3 = "水印操作"
        tabview01_title_n = "使用说明"
        self.tabview01.add(tabview01_title1)
        self.tabview01.add(tabview01_title2)
        self.tabview01.add(tabview01_title3)
        self.tabview01.add(tabview01_title_n)

        # 第一个选项卡 PDF文本比较 fc
        # 第一个选项卡 PDF文本比较 fc
        # 第一个选项卡 PDF文本比较 fc
        # 模式选择
        self.fc_mode_var = customtkinter.StringVar()
        self.fc_mode_var.set("mode_auto")
        self.fc_radio_mode_1 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title1), text="自动模式", variable=self.fc_mode_var, value="mode_auto", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.fc_radio_mode_1.place(x=30, y=10)
        self.fc_radio_mode_2 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title1), text="文本模式", variable=self.fc_mode_var, value="mode_text", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.fc_radio_mode_2.place(x=30, y=40)
        self.fc_radio_mode_3 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title1), text="图片模式", variable=self.fc_mode_var, value="mode_image", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.fc_radio_mode_3.place(x=30, y=70)

        fc_label_mode_radio1 = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title1), text="根据PDF每页内容自动识别。若页中有图片，则此页用图片模式处理；若页中只有文本，则用文本模式处理。", font=self.textbox_font).place(x=130, y=8)
        fc_label_mode_radio2 = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title1), text="运行速度快。只识别文本，不识别图片。如果是图片生成的PDF，那么此模式将识别不到内容。", font=self.textbox_font).place(x=130, y=38)
        fc_label_mode_radio3 = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title1), text="运行速度慢。将PDF全部转成图片，然后进行识别。", font=self.textbox_font).place(x=130, y=68)
        
        # 选择第一个文件
        self.fc_button_select_file_1 = customtkinter.CTkButton(self.tabview01.tab(tabview01_title1), text="请选择第一个文件", command=lambda i=1: self.select_folder(i), width=140, font=self.button_font)
        self.fc_button_select_file_1.place(x=30, y=115)

        self.fc_entry_select_file_1 = customtkinter.CTkEntry(self.tabview01.tab(tabview01_title1), fg_color="#E9EBFE", font=self.entry_font, width=525, height=25, corner_radius=0, border_width=1)
        self.fc_entry_select_file_1.place(x=195, y=115)  

        # 选择第二个文件
        self.fc_button_select_file_2 = customtkinter.CTkButton(self.tabview01.tab(tabview01_title1), text="请选择第二个文件", command=lambda i=2: self.select_folder(i), width=140, font=self.button_font)
        self.fc_button_select_file_2.place(x=30, y=160)

        self.fc_entry_select_file_2 = customtkinter.CTkEntry(self.tabview01.tab(tabview01_title1), fg_color="#E9EBFE", font=self.entry_font, width=525, height=25, corner_radius=0, border_width=1)
        self.fc_entry_select_file_2.place(x=195, y=160)

        # 进行比较
        self.fc_button_compare_file = customtkinter.CTkButton(self.tabview01.tab(tabview01_title1), text="文本比较", command=lambda: threading.Thread(target=self.compare_and_create, args=(self.fc_entry_select_file_1.get(), self.fc_entry_select_file_2.get(),)).start(), width=120, font=self.button_font)
        self.fc_button_compare_file.place(x=150, y=205)

        # log打印
        self.fc_textbox_log = customtkinter.CTkTextbox(self.tabview01.tab(tabview01_title1), bg_color="#000001",width=690, height=180, state=customtkinter.DISABLED, font=self.textbox_font, corner_radius=0)
        self.fc_textbox_log.place(x=30, y=250)
        
        
        # 第二个选项卡 文本转换  ft
        # 第二个选项卡 文本转换  ft
        # 第二个选项卡 文本转换  ft
        # 文件选择
        self.ft_button_select_file = customtkinter.CTkButton(self.tabview01.tab(tabview01_title2), text="请选择要转换的文件", command=lambda i=1: self.select_convert_file(i), width=140, font=self.button_font)
        self.ft_button_select_file.place(x=30, y=10)

        self.ft_entry_select_file = customtkinter.CTkEntry(self.tabview01.tab(tabview01_title2), fg_color="#E9EBFE", font=self.entry_font, width=525, height=25, corner_radius=0, border_width=1)
        self.ft_entry_select_file.place(x=195, y=12)
        
        # 模式选择
        self.ft_mode_var = customtkinter.StringVar()
        self.ft_mode_var.set("word2pdf")
        self.ft_radio_mode_1 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title2), text="word 转 pdf", variable=self.ft_mode_var, value="word2pdf", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.ft_radio_mode_1.place(x=30, y=70)

        # 文件转换
        self.ft_button_convert = customtkinter.CTkButton(self.tabview01.tab(tabview01_title2), text="文本转换", command=lambda: threading.Thread(target=self.file_convert, args=(self.ft_entry_select_file.get(),)).start(), width=120, font=self.button_font)
        self.ft_button_convert.place(x=30, y=205)
        
        
        # 第三个选项卡 水印操作  wh
        # 第三个选项卡 水印操作  wh
        # 第三个选项卡 水印操作  wh
        # 文件选择
        self.wh_button_select_file = customtkinter.CTkButton(self.tabview01.tab(tabview01_title3), text="请选择要处理的文件", command=lambda i=1: self.wh_select_convert_file(i), width=140, font=self.button_font)
        self.wh_button_select_file.place(x=30, y=10)

        self.wh_entry_select_file = customtkinter.CTkEntry(self.tabview01.tab(tabview01_title3), fg_color="#E9EBFE", font=self.entry_font, width=525, height=25, corner_radius=0, border_width=1)
        self.wh_entry_select_file.place(x=195, y=12)
        
        # 处理模式选择
        self.wh_mode_add_del_var = customtkinter.StringVar()
        self.wh_mode_add_del_var.set("add")
        self.wh_radio_add_del_mode_1 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title3), text="添加水印", variable=self.wh_mode_add_del_var, value="add", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.wh_radio_add_del_mode_1.place(x=30, y=70)
        self.wh_radio_add_del_mode_2 = customtkinter.CTkRadioButton(self.tabview01.tab(tabview01_title3), text="删除水印", variable=self.wh_mode_add_del_var, value="del", border_color="gray", font=self.radio_font, radiobutton_width=18, radiobutton_height=18, border_width_checked=5)
        self.wh_radio_add_del_mode_2.place(x=30, y=95)
        
        # 其他
        # 水印内容
        wh_label_w_c = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印内容", font=self.entry_font).place(x=30, y=138)
        self.wh_entry_w_c = customtkinter.CTkEntry(self.tabview01.tab(tabview01_title3), fg_color="#E9EBFE", font=self.entry_font, width=300, height=25, corner_radius=0, border_width=1)
        self.wh_entry_w_c.place(x=100, y=140)
        
        # 字体类型
        wh_label_font_type = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="字体类型", font=self.entry_font).place(x=30, y=173)
        self.wh_combobox_font_type = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=self.pdf_font_list, state="readonly", width=150, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_font_type.place(x=100, y=175)
        self.wh_combobox_font_type.set(self.pdf_font_list[0])

        # 字体大小
        wh_label_font_size = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="字体大小", font=self.entry_font).place(x=30, y=208)
        self.wh_combobox_font_size = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ str(i) for i in range(15, 51, 5)], width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_font_size.place(x=100, y=210)
        self.wh_combobox_font_size.set("40")

        # 水印行数
        wh_label_w_row = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印行数", font=self.entry_font).place(x=30, y=243)
        self.wh_combobox_w_row = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ str(i) for i in range(1, 6, 1)], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_row.place(x=100, y=245)
        self.wh_combobox_w_row.set("1")
        
        # 水印列数
        wh_label_w_col = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印列数", font=self.entry_font).place(x=30, y=278)
        self.wh_combobox_w_col = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ str(i) for i in range(1, 6, 1)], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_col.place(x=100, y=280)
        self.wh_combobox_w_col.set("1")
        
        # 水印旋转角度
        wh_label_w_angle = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印角度", font=self.entry_font).place(x=30, y=313)
        self.wh_combobox_w_angle = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ "15", "30", "45", "60", "75", "105", "120", "135", "150", "165", ], width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_angle.place(x=100, y=315)
        self.wh_combobox_w_angle.set("45")
        
        # 水印偏移位置 x
        wh_label_w_offset_x = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印偏移x", font=self.entry_font).place(x=240, y=243)
        self.wh_combobox_w_offset_x = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ str(i) for i in range(0, 21, 2)], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_offset_x.place(x=320, y=245)
        self.wh_combobox_w_offset_x.set("0")
        
        # 水印偏移位置 y
        wh_label_w_offset_y = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印偏移y", font=self.entry_font).place(x=240, y=278)
        self.wh_combobox_w_offset_y = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ str(i) for i in range(0, 21, 2)], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_offset_y.place(x=320, y=280)
        self.wh_combobox_w_offset_y.set("0")
        
        # 水印透明度
        wh_label_w_transparency = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="透明度", font=self.entry_font).place(x=250, y=208)
        self.wh_combobox_w_transparency = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ "0.2", "0.3", "0.4", "0.5"], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_transparency.place(x=320, y=210)
        self.wh_combobox_w_transparency.set("0.2")
        
        # 水印阈值
        wh_label_w_threshold = customtkinter.CTkLabel(self.tabview01.tab(tabview01_title3), text="水印阈值", font=self.entry_font).place(x=245, y=313)
        self.wh_combobox_w_threshold = customtkinter.CTkComboBox(self.tabview01.tab(tabview01_title3), values=[ "400", "500", "550", "600", "650"], state="readonly", width=90, height=20, font=self.combox_font, corner_radius=0)
        self.wh_combobox_w_threshold.place(x=320, y=315)
        self.wh_combobox_w_threshold.set("600")

        # 说明窗口
        wh_readme = "\n" + \
                    "说明:\n\n" + \
                    "1 水印偏移x/y可以调整水印的位置。\n\n" + \
                    "2 本地字体类型的路径 'C:\Windows\Fonts'，字体类型通过 config.py 进行配置。\n\n" + \
                    "3 生成的pdf文件在所选择的pdf文件目录中。\n\n" + \
                    "4 水印如果没有去除，可适当降低水印阈值。\n\n"

        self.wh_textbox_readme = customtkinter.CTkTextbox(self.tabview01.tab(tabview01_title3), width=280, height=368, corner_radius=0)
        self.wh_textbox_readme.place(x=440, y=60)
        self.wh_textbox_readme.insert("0.0", wh_readme)
        self.wh_textbox_readme.configure(state="disabled")

        # 文件处理
        self.wh_button_pdf = customtkinter.CTkButton(self.tabview01.tab(tabview01_title3), text="文件处理", command=lambda: threading.Thread(target=self.pdf_create_del_watermark).start(), width=140, font=self.button_font)
        self.wh_button_pdf.place(x=100, y=405)


        

        # 第N个选项卡 使用说明 rd
        # 第N个选项卡 使用说明 rd
        # 第N个选项卡 使用说明 rd
        rd_str = '''
                 1 程序运行目录中不能有中文字符。\n
                 2 此程序只在本机运行，不会向其他设备传递任何信息。\n
                 3 此程序所生成的文件结果仅供参考。\n
                 '''
        self.rd_textbox = customtkinter.CTkTextbox(self.tabview01.tab(tabview01_title_n), width=560, height=400, corner_radius=0)
        self.rd_textbox.place(x=45, y=20)
        self.rd_textbox.insert("0.0", rd_str)  # insert at line 0 character 0
        self.rd_textbox.configure(state="disabled")  # configure textbox to be read-only


        
        
app = App()
app.protocol("WM_DELETE_WINDOW", app.on_closing)
app.mainloop()

