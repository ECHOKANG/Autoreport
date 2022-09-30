#-*- coding: utf-8 -*-
from email.mime import image
import tkinter as tk
import openpyxl
import fitz
import os
import sys
from PIL import Image, ImageDraw, ImageFont
# e:\pytest\autoreport\autoreport_main.py
BASE_DIR = os.path.dirname(os.path.realpath(sys.argv[0]))
#读取文件根目录
#print(BASE_DIR)

wind = tk.Tk()
wind.title('文档生成器测试版')
wind.geometry('450x800')
'''
tk.Label(wind,text='Excel文件路径').place(x=50,y=150)
tk.Label(wind,text='图片所在路径').place(x=50,y=200)
tk.Label(wind,text='保存路径').place(x=50,y=250)
tk.Label(wind,text='文件目录示例').place(x=50,y=350)
'''
tk.Label(wind,text='加载人数').place(x=50,y=300)
tk.Label(wind,text=str(BASE_DIR)).place(x=50,y=375)
# entry_local_excel = tk.Entry(wind)
# entry_local_excel.place(x=160,y=150)
# entry_local_pic = tk.Entry(wind)
# entry_local_pic.place(x=160,y=200)
# entry_local_save = tk.Entry(wind)
# entry_local_save.place(x=160,y=250)
entry_codeline = tk.Entry(wind)
entry_codeline.place(x=160,y=300)

print('//////')

def create_pdf():
    path_excel = BASE_DIR + '\\1register.xlsx' #初始化图片存储位置
    path_image = BASE_DIR + '\\test.png'#初始化excel文件位置
    path_model = BASE_DIR + '\\model.jpg'#初始化个人档案模板
    
    line_name = entry_codeline.get()
    linesum = int(line_name) + 1
    for i in range(2,linesum):
        line_code = str(i)
        

        print(line_code)
        print('初始化完成')
        martic = 'A' + line_code
        name = 'B' + line_code
        gender = 'C' + line_code
        hometown = 'D'+ line_code
        sexhis = 'E' + line_code
        brithday = 'F' + line_code
        height = 'G' + line_code
        weight = 'H' + line_code
        edu = 'I' + line_code
        fam = 'J' + line_code
        worloca = 'K' + line_code
        pers_des = 'L' + line_code
        photo = 'M' + line_code
        keydes = 'N' + line_code
        recomer = 'O' + line_code
        wb = openpyxl.load_workbook(path_excel) #path_excel+code
        sheet = wb['index']
        sheetname = wb.sheetnames
        print(sheetname)
        shall_a = sheet[str(martic)]
        shall_b = sheet[str(name)]
        shall_c = sheet[str(gender)]
        shall_d = sheet[str(hometown)]
        shall_e = sheet[str(sexhis)]
        shall_f = sheet[str(brithday)]
        shall_g = sheet[str(height)]
        shall_h = sheet[str(weight)]
        shall_i = sheet[str(edu)]
        shall_j = sheet[str(fam)]
        shall_k = sheet[str(worloca)]
        shall_l = sheet[str(pers_des)]
        shall_m = sheet[str(photo)]
        shall_n = sheet[str(keydes)]
        shall_o = sheet[str(recomer)]

        #存储表格内容
        values_a = shall_a.value
        values_b = shall_b.value
        values_c = shall_c.value
        values_d = shall_d.value
        values_e = shall_e.value
        values_f = shall_f.value
        values_g = shall_g.value
        values_h = shall_h.value
        values_i = shall_i.value
        values_j = shall_j.value
        values_k = shall_k.value
        values_l = shall_l.value
        values_m = shall_m.value
        values_n = shall_n.value
        values_o = shall_o.value
        info_list = [values_b, #BIOLIST
        values_c,
        values_d,
        values_e,
        values_f,
        values_g,
        values_h,
        values_i,
        values_j,
        values_k,
        values_l,
        values_m,
        values_n,
        values_o,
        values_a]

        print('数据获取完成')
        print(values_a)
        #创建PDF文档i
        print('开始创建文档')
        filerename = line_code + 'docpic.jpg'
        save_path = BASE_DIR + '\\' + filerename
        print(save_path)
        def poster(path, str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, mar=str(info_list[14])):
            '''
            path:海报模板文件路径
            添加的文字
            '''
            png1 = Image.open(path) #打开文件
            draw = ImageDraw.Draw(png1)
            font_name = ImageFont.truetype('E:\\pytest\\autoreport\\SimSun.ttf', 80, encoding="utf-8")  # 设置字体
            font_main = ImageFont.truetype('E:\\pytest\\autoreport\\SimSun.ttf', 60, encoding="utf-8")
            draw.text((1120, 1740), str1, font=font_name, fill='black', Literal='center')  # name list[0]
            draw.text((524, 2169), str2, font=font_main, fill='black', Literal='center')    # gender list[1]
            draw.text((524, 2267), str3, font=font_main, fill='black', Literal='center')    # hometown list[2]
            draw.text((654, 2371), str4, font=font_main, fill='black', Literal='center')    # sexhis list[3]
            draw.text((654, 2470), str5, font=font_main, fill='black', Literal='center')    # brithday list[4]
            draw.text((654, 2604), str6, font=font_main, fill='black', Literal='center')    # fam list[8]
            draw.text((1519, 2169), str7, font=font_main, fill='black', Literal='center')   # high list[5]
            draw.text((1519, 2267), str8, font=font_main, fill='black', Literal='center')   # weight list[6]
            draw.text((1519, 2370), str9, font=font_main, fill='black', Literal='center')   # edu list[7]
            draw.text((1650, 2470), str10, font=font_main, fill='black', Literal='center')  # workplace list[9]
            draw.text((524, 2827), str11, font=font_main, fill='black', Literal='center')
            draw.text((1111, 467), 'SF'+ mar +'V', font=font_name, fill='black', Literal='center') #martic no
            user_img = str(BASE_DIR + '//' + str(info_list[14]) + '.jpg')  # 海报名称 
            print((user_img))
            imgfile = Image.open(user_img)
            resized = imgfile.resize((612,801))
            png1.paste(resized,(935,771))
            '''
            进行中，需要配置添加图片
            1.设置弹窗提示完成
            2.GUI多余内容删除
            3.设置性格
            '''
            png1.save(save_path)  # 保存海报
        poster(path_model,str1=str(info_list[0]),str2=str(info_list[1]),str3=str(info_list[2]),str4=str(info_list[3]),str5=str(info_list[4]),str6=str(info_list[8]),str7=str(info_list[5]),str8=str(info_list[6]),str9=str(info_list[7]),str10=str(info_list[9]),str11=info_list[10])
        print('文档创建完成，请查看' + save_path)

'''
    # pdfmetrics.registerFont(TTFont('SimSun', 'SimSun.ttf')) #加载字体文件
    # doc = SimpleDocTemplate(pdffile)
    # styles = getSampleStyleSheet()
    # styles.add(ParagraphStyle(fontName='SimSun', name='Song_title', leading=40, fontSize=22))  #自己增加新注册的字体
    # styles.add(ParagraphStyle(fontName='SimSun', name='Song', leading=20, fontSize=12))

    Title = "客户档案"
    pageinfo = "客户信息详情"

    style_body = styles['Song']
    stylepic = styles['Title']
    body = []#主容器


    body.append(Paragraph(info_list[0],styles['Song_title']))
    body.append(Paragraph('年龄'+':'+str(info_list[4]) ,style_body))
    body.append(Paragraph('祖籍'+':'+info_list[2] ,style_body))
    body.append(Paragraph('婚姻状况'+':'+info_list[3] ,style_body))
    body.append(Paragraph('身高'+':'+info_list[5] ,style_body))
    body.append(Paragraph('体重'+':'+str(info_list[6]),style_body))
    body.append(Paragraph('学历'+':'+info_list[7],style_body))
    body.append(Paragraph('家庭状况'+':'+info_list[8],style_body))
    body.append(Paragraph('工作地点'+':'+info_list[9],style_body))
    body.append(Paragraph('性格'+':'+info_list[10],style_body))
    #body.append(Paragraph(''+':'+info_list[],style_body))
    #body.append(Paragraph(''+':'+info_list[],style_body))
    doc.build(body)

    input_file = pdffile
    output_file = save_path + line_code + ".pdf"
    adding_image_file = path_image
    # define the position (upper-right corner)
    image_rectangle = fitz.Rect(350,10,550,150)

    # retrieve the first page of the PDF
    file_handle = fitz.open(input_file)
    first_page = file_handle[0]

    # add the image
    first_page.insert_image(image_rectangle, filename=adding_image_file)

    file_handle.save(output_file)


'''       
butt_submit = tk.Button(wind,text='生成',command=create_pdf)
butt_submit.place(x=350,y=400)
wind.mainloop()