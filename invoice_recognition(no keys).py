# -*- coding: utf-8 -*-
"""
Created on Thu Mar 12 10:20:44 2020

@author: admin
"""

from pandas import DataFrame,ExcelWriter,to_datetime
import os
import re
from json import loads 
from base64 import b64encode
from numpy import zeros
import fitz
import datetime
import time

##导入腾讯AI api
from tencentcloud.common import credential
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.ocr.v20181119 import ocr_client, models

from tkinter import filedialog,Tk,Label,Button,Menu,Entry,ttk

#定义函数,来自于官方文档
def excelFromPictures(path,picture):
    SecretId = ""
    SecretKey = ""     
    
    with open(picture,"rb") as f:
            img_data = f.read()
    img_base64 = b64encode(img_data)
    cred = credential.Credential(SecretId, SecretKey)  #ID和Secret从腾讯云申请
    httpProfile = HttpProfile()
    httpProfile.endpoint = "ocr.tencentcloudapi.com"

    clientProfile = ClientProfile()
    clientProfile.httpProfile = httpProfile
    client = ocr_client.OcrClient(cred, "ap-shanghai", clientProfile)

    req = models.VatInvoiceOCRRequest()
    params = '{"ImageBase64":"' + str(img_base64, 'utf-8') + '"}'
    req.from_json_string(params)
#    false=0
    try:

        resp = client.VatInvoiceOCR(req) 
        #     print(resp.to_json_string())

    except TencentCloudSDKException as err:
        print("识别",picture,"错误[",err,"]\n可重试")
        


    ##提取识别出的数据，并且生成json
    result1 = loads(resp.to_json_string())

#    print(result1)
#    print(resp.to_json_string())
    
    invoicedf = DataFrame(zeros(5).reshape(1,5),
                             columns=["发票代码","发票号码","开票日期","合计金额","小写金额"])
    for item in result1['VatInvoiceInfos']:
        if item["Name"] in ["发票代码","发票号码","开票日期","合计金额","小写金额"]:
           invoicedf[item["Name"]] = item["Value"]
    

#
#    writer = ExcelWriter(path+"/tables/" +re.match(".*\.",f.name).group()+"xlsx", engine='xlsxwriter')
#    data.to_excel(writer,sheet_name = 'Sheet1', index=False,header = False)
#    writer.save()
#    
    print("已经完成[" + f.name + "]的识别")
    return invoicedf



def pyMuPDF_fitz():
    '''
    pdf转图片
    '''
    #picture_path = input("请输入表格图片路径：")
    pdfPath = entry_filename1.get()
    path = os.path.dirname(pdfPath)
    os.chdir(path)
    f_name = os.path.basename(pdfPath).split(".")[0]
    imagePath = os.path.join(path,f_name+"-images")       

    startTime_pdf2img = datetime.datetime.now()#开始时间
    
    angle = angleChosen.get()
    zoom_num  = zoomCoef.get()
#    print("imagePath="+imagePath)
    pdfDoc = fitz.open(pdfPath)
    for pg in range(pdfDoc.pageCount):
        page = pdfDoc[pg]
        rotate = int(angle)
        # 每个尺寸的缩放系数为1.3，这将为我们生成分辨率提高2.6的图像。

        zoom_x = zoom_num #(1.33333333-->1056x816)   (2-->1584x1224) #数字越大清晰度越高，但图片太大又无法上传至腾讯云
        zoom_y = zoom_num
        mat = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
#        mat = fitz.Matrix().preRotate(rotate)
        pix = page.getPixmap(matrix=mat, alpha=False)
        
        if not os.path.exists(imagePath):#判断存放图片的文件夹是否存在
            os.makedirs(imagePath) # 若图片文件夹不存在就创建
        
        pix.writePNG(imagePath+'/'+f_name+'-images_%s.png' % pg)#将图片写入指定的文件夹内
        
    endTime_pdf2img = datetime.datetime.now()#结束时间
    print('pdf已转换，用时%ds'%(endTime_pdf2img - startTime_pdf2img).seconds)
    
    

def batch():
    '''
    发票图片识别
    '''
    file_str = entry_filename2.get()
#    print(file_str)
    file_names = re.split(r"[{} ]",file_str)
#    print(file_names)
    file_names = [f.lstrip() for f in file_names if f not in [""," "]]
    file_names = [f.rstrip() for f in file_names]
    pictures_path = os.path.dirname(file_names[0])
    path = os.path.dirname(pictures_path)
    os.chdir(pictures_path)
    
    f_name = pictures_path.split("/")[-1].split("-")[0]
#    print(f_name)
    
    
    pictures = [os.path.basename(f) for f in file_names]
    
    table_path = os.path.join(path,"tables")
    
    if not os.path.exists(table_path):
        os.mkdir(table_path)
    
    invoice_table = DataFrame(columns=["发票代码","发票号码","开票日期","合计金额","小写金额"])
    for pic in pictures:
        try:
            invoicedf= excelFromPictures(path,pic)
            invoice_table = invoice_table.append(invoicedf)
            time.sleep(1)
        except:
            pass

        
        
    invoice_table = invoice_table.reset_index(drop=True)
    invoice_table['发票号码'] = invoice_table['发票号码'].str[2:]
    invoice_table['合计金额'] = invoice_table['合计金额'].str[1:]
    invoice_table['小写金额'] = invoice_table['小写金额'].str[1:]
    invoice_table['开票日期'] = to_datetime(invoice_table['开票日期'],format = '%Y年%m月%d日').apply(lambda x : x.strftime('%Y%m%d'))
    invoice_table.rename(columns={'开票日期':'发票日期', '合计金额':'税前金额',"小写金额":"价税合计"}, inplace = True)
    invoice_table["税前金额"] = invoice_table["税前金额"].astype(float)#此处增加保存两位小数
    invoice_table["价税合计"] = invoice_table["价税合计"].astype(float)
    
    writer = ExcelWriter(path+"/tables/"+f_name+"-invoice_df.xlsx", engine='xlsxwriter')
    invoice_table.to_excel(writer,sheet_name = 'Sheet1', index=True,header=True)
    writer.save()     
    
    print("已完成发票识别，请打开"+path+"/tables查看")

window = Tk()
window.title('发票识别神器')  
window.geometry('700x200')

def file_input_one():
    filename = filedialog.askopenfilename(title='导入图片文件')
    entry_filename1.insert('insert', filename) 

def file_input_batch():
    filename = filedialog.askopenfilenames(title='导入图片文件')
    entry_filename2.insert('insert', filename) 
    
menubar = Menu(window)
filemenu = Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=filemenu)
filemenu.add_command(label='Open_pdf', command=file_input_one)
filemenu.add_command(label='Open_images', command=file_input_batch)
window.config(menu=menubar)



l1 = Label(window, text="pdf转图片",font=("宋体", 10, 'bold'))
l1.grid(column=0, row=0)

entry_filename1 = Entry(window, width=30,font=("arial", 10))
entry_filename1.grid(column=0, row=1)


l3 = Label(window, text="旋转角度",font=("宋体", 10, 'bold'))
l3.grid(column=1, row=0) 


angleChosen = ttk.Combobox(window, width=4) #3
angleChosen['values'] = (0,90,180,270) # 4
angleChosen.grid(column=1, row=1) # 5
angleChosen.current(0) # 6


l4 = Label(window, text="缩放系数",font=("宋体", 10, 'bold'))
l4.grid(column=2, row=0) 
zoomCoef = Entry(window, width=4,font=("arial", 10))
zoomCoef.grid(column=2, row=1) # 5



b1 = Button(window, text="开始转换",command=pyMuPDF_fitz)
b1.grid(column=3, row=1)



l2 = Label(window, text="发票图片识别",font=("宋体", 10, 'bold'))
l2.grid(column=0, row=2)

entry_filename2 = Entry(window, width=30,font=("arial", 10))
entry_filename2.grid(column=0, row=3)

def test_batch():
    file_str = entry_filename2.get()
    print(file_str)
    file_names = re.split(r"[{} ]",file_str)
    print(file_names)

b2 = Button(window, text="开始识别",command=batch)
b2.grid(column=3, row=3)

tips = Label(window, text="注：图片名称中不允许有空格；缩放系书越大图片清晰度越高，但大小不要超过700k",font=("仿宋", 8))
tips.grid(column=0,row=4)

window.mainloop()