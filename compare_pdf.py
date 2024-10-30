import os
import sys
import shutil
import threading
from itertools import product
import fitz
from paddleocr import PaddleOCR
from difflib import HtmlDiff
import customtkinter
from CTkMessagebox import CTkMessagebox





class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        customtkinter.deactivate_automatic_dpi_awareness()
        self.textbox_font = customtkinter.CTkFont(family="微软雅黑", size=12)
        self.entry_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.button_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.msg_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.radio_font = customtkinter.CTkFont(family="微软雅黑", size=14)
        self.font = customtkinter.CTkFont(family="微软雅黑", size=22)
        self.label_font = customtkinter.CTkFont(family="Arial", size=24)
        self.title("PDF文本比较工具 v1.2.0")
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
        
        self.user_ui()
        

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
        
        #return total_res
        
        
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


    # 单张图片去水印
    def mt_pic_remove_watermark(self, index, page, pdf_file, img_folder):
            rotate = int(0)
            # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高4倍的图像
            # zoom_x, zoom_y = 2, 2
            zoom_x, zoom_y = 2, 2
            trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
            pixmap = page.get_pixmap(matrix=trans, alpha=False)
            for pos in product(range(pixmap.width), range(pixmap.height)):
                rgb = pixmap.pixel(pos[0], pos[1])
                if (sum(rgb) >= 600):
                    pixmap.set_pixel(pos[0], pos[1], (255, 255, 255))
            pixmap.pil_save(os.path.join(img_folder, "pdf_split_" + str(index+1).zfill(3) + ".png"))
            print(f"    {pdf_file} 第 {index+1} 页水印去除完成")
            self.status_message_add(f"        {pdf_file} 第 {index+1} 页去水印完成")
            
            
    # pdf 每一页转图片, 然后图片去水印
    def pdf_to_pic_remove_watermark(self, pdf_file, img_folder):
        t_list = []
        pdf = fitz.open(pdf_file)
        
        for index, page in enumerate(pdf):
            self.mt_pic_remove_watermark(index, page, pdf_file, img_folder)
        
        
    # 获取 文本文件内容
    def get_file_content(self, file_path):
        lines = []
        with open(file_path, mode="r", encoding="utf-8") as f:
            lines = f.read().splitlines()
        return lines 
    
    
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


    # 状态栏打印信息 初始化
    def status_message_init(self, message):
        self.fc_textbox_log.configure(state=customtkinter.NORMAL)
        self.fc_textbox_log.delete(1.0, customtkinter.END)
        self.fc_textbox_log.insert(customtkinter.END, message + "\n")
        self.fc_textbox_log.configure(state=customtkinter.DISABLED)
        
        
    # 状态栏打印信息 增加
    def status_message_add(self, message):
        self.fc_textbox_log.configure(state=customtkinter.NORMAL)
        self.fc_textbox_log.insert(customtkinter.END, message + "\n")
        self.fc_textbox_log.configure(state=customtkinter.DISABLED)
        self.fc_textbox_log.see(customtkinter.END)
        
        
    def on_closing(self):
            self.destroy()
            self.quit()
            sys.exit()


    # UI 页面
    def user_ui(self):
        # 标签页
        customtkinter.CTkTabview._outer_button_overhang = 15
        self.tabview01 = customtkinter.CTkTabview(master=self, width=755, height=505, fg_color="#FDE6E0", segmented_button_fg_color="#FDE6E0", segmented_button_selected_color="#4F86EC", segmented_button_unselected_color="gray", border_width=2, corner_radius=0, bg_color="#FDE6E0", segmented_button_font=self.button_font)
        self.tabview01.place(x=-2, y=5)

        # 添加2个选项卡
        tabview01_title1 = "PDF文本比较"
        tabview01_title2 = "使用说明"
        self.tabview01.add(tabview01_title1)
        self.tabview01.add(tabview01_title2)

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

        # 第二个选项卡 使用说明 rd
        # 第二个选项卡 使用说明 rd
        # 第二个选项卡 使用说明 rd
        rd_str = '''
                 1 程序运行目录中不能有中文字符。\n
                 2 此程序只在本机运行，不会向其他设备传递任何信息。\n
                 3 目前只进行pdf文件的文本内容比对，不进行格式的比对。\n
                 4 此程序不能完全正确识别pdf文件内容，比较内容仅供参考。\n
                 '''
        self.rd_textbox = customtkinter.CTkTextbox(self.tabview01.tab(tabview01_title2), width=560, height=400, corner_radius=0)
        self.rd_textbox.place(x=45, y=20)
        self.rd_textbox.insert("0.0", rd_str)  # insert at line 0 character 0
        self.rd_textbox.configure(state="disabled")  # configure textbox to be read-only


        
        
app = App()
app.protocol("WM_DELETE_WINDOW", app.on_closing)
app.mainloop()

