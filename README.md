## 软件说明
用于pdf内容比较，比较结果不保证完全正确，仅供参考。  
  
**软件运行模式：自动模式、文本模式、图片模式。**  
文本模式：运行速度快。只识别文本，不识别图片。如果是图片生成的PDF，那么此模式将识别不到内容。  
图片模式：运行速度慢。将PDF全部转成图片，然后进行识别。  
自动模式：根据PDF每页内容自动识别。若页中有图片，则此页用图片模式处理；若页中只有文本，则用文本模式处理。  

**重点：**  
1 程序运行目录中不能有中文字符。  
2 此程序只在本机运行，不会向其他设备传递任何信息。  
3 目前只进行pdf文件的文本内容比对，不进行格式的比对。  
4 此程序不能完全正确识别pdf文件内容，比较内容仅供参考。  

## 环境说明
已在下面环境中进行验证。  
操作系统：win11  
Python：3.10  

## 安装说明
```py
git clone https://github.com/bibo19842003/compare_pdf
cd compare_pdf
pip install requirement.txt
```
参考 https://github.com/TomSchimansky/CustomTkinter/pull/2605/files 修改 CustomTkinter 组件的相应文件。  

## 代码运行
```py
python compare_pdf.py
```

## 打包EXE文件
**1、运行代码文件**  
```py
python compare_pdf.py
```

**2、执行打包命令**
```py
pyinstaller -w --collect-all paddleocr --collect-all pyclipper --collect-all imghdr --collect-all skimage --collect-all imgaug --collect-all scipy.io --collect-all lmdb  --collect-all requests -y compare_pdf.py
```

**3、添加模型相关文件**  
将 *ch_ppocr_mobile_v2.0_cls_infer*  *ch_PP-OCRv4_det_infer*  *ch_PP-OCRv4_rec_infer* 三个文件夹复制到 dist/compare_pdf/_internal 目录下，若没有那3个目录，执行第一步运行代码文件会从网络下载并生成。  

**4、添加dll文件**  
将python安装包目录 Lib/site-packages/paddle/libs 下面的所有文件拷贝到 dist/compare_pdf/_internal/paddle/libs 目录下。  

**此时，dist\compare_pdf 目录下的 compare_pdf.exe 可以正常运行。**
