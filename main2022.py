from tkinter import *
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import base64
import requests
from lxml import etree
import pandas as pd
from pandas import DataFrame
import threading
import os
import time
from memory_pic import *

def get_pic(pic_code, pic_name):
    image = open(pic_name, 'wb')
    image.write(base64.b64decode(pic_code))
    image.close()

def search():
    print("START ENGINE->")
    filename = 'BASE.xlsx'
    print("->DONE")
    print("->")
    try:
        global sheet
        sheet = pd.read_excel(filename, "Sheet1",converters={'基金代码':str})
        threads=[]
        for row in sheet.index.values:
            var2 = sheet.iloc[row, 0]
            var2 = str(var2).zfill(6)
            t = threading.Thread(target=search_code, args=(var2,row))
            threads.append(t)
        for t in threads:
            if threading.active_count() <= 10:
                t.start()
            if threading.active_count() == 10:
                t.join()
        for t in threads:
            if t.is_alive():
                t.join()
    except Exception as e:
        print(e)
    finally:
        print("->")
        print("CLOSE ENGINE->")
        DataFrame(sheet).to_excel(filename, sheet_name='Sheet1', index=False, header=True)
        print("->DONE")
        print("-----------------------------")
        print("MISSION COMPLETE")

def search_code(var1,row):
    ARR2={}
    ARR2[0]=var1
    url = "https://fund.eastmoney.com/" + ARR2[0] + ".html"
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 Safari/537.36'
    }
    page = requests.get(url, headers=headers)
    try:
        html = etree.HTML(page.content.decode('utf-8'))
        # 基金名称
        #ARR1[1] = html.xpath("//div[@class='fundDetail-tit']")[0].xpath('string(.)') #名称+代码
        ARR2[1] = html.xpath("//div[@class='fundDetail-tit']/div")[0].text #名称
        # 单位净值
        ARR2[2] = html.xpath("//dl[@class='dataItem02']/dd/span")[0].xpath('string(.)')
        # 累计净值
        ARR2[3] = html.xpath("//dl[@class='dataItem03']/dd/span")[0].xpath('string(.)')
        # 成立来
        ARR2[7] = html.xpath("//dl[@class='dataItem03']/dd[3]/span[2]")[0].xpath('string(.)')
        # 基金规模
        ARR2[4] = html.xpath("//div[@class='infoOfFund']/table/tr[1]/td[2]")[0].xpath('string(.)')
        ARR2[4] =ARR2[4].replace("基金规模：", "").replace("）", "").replace("（", "|")
        # 基金经理
        ARR2[5] = html.xpath("//div[@class='infoOfFund']/table/tr[1]/td[3]")[0].xpath('string(.)')
        ARR2[5] = ARR2[5].replace("基金经理：", "")
        # 成立日
        ARR2[6] = html.xpath("//div[@class='infoOfFund']/table/tr[2]/td[1]")[0].xpath('string(.)')
        ARR2[6] = ARR2[6].replace("成 立 日：", "")
        # 阶段涨跌幅
        ARR2[8] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[2]")[0].xpath('string(.)')
        ARR2[9] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[3]")[0].xpath('string(.)')
        ARR2[10] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[4]")[0].xpath('string(.)')
        ARR2[11] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[5]")[0].xpath('string(.)')
        ARR2[12] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[6]")[0].xpath('string(.)')
        ARR2[13] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[7]")[0].xpath('string(.)')
        ARR2[14] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[8]")[0].xpath('string(.)')
        ARR2[15] = html.xpath("//li[@id='increaseAmount_stage']/table/tr[2]/td[9]")[0].xpath('string(.)')
        sys.stdout.write('%s %s  %s \n' % (ARR2[0],ARR2[2],url))
    except:
        sys.stdout.write('%s 基金异常  请核对基金代码或状态 \n' %(ARR2[0]))
    finally:
        if len(ARR2) == 16:
            global sheet
            for ii in range(1, 16):
                sheet.iloc[row, ii] = ARR2[ii]
                ii += 1
    return ARR2

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")
        self.widget.see(tk.END)

class MyLabel(Label):
    def __init__(self, master, filename):
        im = Image.open(filename)
        seq =  []
        try:
            while TRUE:
                seq.append(im.copy())
                im.seek(len(seq))
        except EOFError:
            pass

        try:
            self.delay = im.info['duration']
        except KeyError:
            self.delay = 100
        first = seq[0].convert('RGBA')
        self.frames = [ImageTk.PhotoImage(first)]
        Label.__init__(self, master, image=self.frames[0])
        temp = seq[0]
        for image in seq[1:]:
            temp.paste(image)
            frame = temp.convert('RGBA')
            self.frames.append(ImageTk.PhotoImage(frame))
        self.idx = 0
        self.after(self.delay*2, self.play)
        self.num = 1
    def cancel(self):
        self.num=1
    def play(self):
        if self.num:
            self.config(image=self.frames[self.idx])
            self.idx += 1
            if self.idx == len(self.frames):
                self.idx = 0
            self.after(self.delay*2, self.play)

def s_play():
    if anim.num==2:
        anim.config(image=anim.frames[anim.idx])
        anim.idx += 1
        if anim.idx == len(anim.frames):
            anim.idx = 0
        anim.after(anim.delay, s_play)

def get_time():
    time1 = ''
    time2 = time.strftime('%Y-%m-%d %H:%M:%S')
    if time2 != time1:
        clock.configure(text=str(time2))
        clock.after(200, get_time)

def run():
    thread1 = threading.Thread(target=search)
    thread1.setDaemon(True)
    thread1.start()
    thread2 = threading.Thread(target=check,args=(thread1,))
    thread2.setDaemon(True)
    thread2.start()

def check(thread_var1):
    flag=0
    while True:
        if thread_var1.is_alive():
            btn1['state'] = tk.DISABLED
            btn1['text'] = 'RUNNING'
            if flag<1:
                anim.num = 2
                s_play()
                flag=1
        else:
            btn1['state'] = tk.NORMAL
            btn1['text'] = 'RUN'
            anim.cancel()
            break
        root.update()

def sp(event):
    anim.num=0
    print("双击LOGO,加载自定义基金代码EXCEL文件 ->")
    my_filetypes = [('all files', '.*'), ('text files', '.txt')]
    global filename
    filename = filedialog.askopenfilename(parent=root,initialdir=os.getcwd(),
                                          title="请选择你要处理的EXCEL",filetypes=my_filetypes)
    anim.num = 1
    anim.play()
    print(filename)
    print("->")
    print("->")
    print("-> 加载完成")
    print("->")

filename='BASE.xlsx'
root = Tk()
root.title('公募基金净值爬虫 V4.1')
root.config(bg='black',bd = 0,relief="sunken")
get_pic(bg_gif, 'bg_gif')
anim = MyLabel(root, "bg_gif")
anim.config(bg='black',bd = 0)
root.resizable(width=False, height=False)
anim.bind("<Double-Button-1>", sp)
anim.pack()

text=Text(root,width = 55,height=20,bg='black',foreground = 'white',relief="solid",insertbackground='white')
text.tag_configure("stderr")
sys.stdout = TextRedirector(text, "stdout")
text.pack()

btn1=Button(root, text='RUN',width=7,font= ("Lucida Sans", 15),command=run)
btn1.pack()
clock = Label(root, text="", height=1,bg='black',fg='white')
clock.pack()
get_time()
Label(root,text=' Powered by LEO    |    RAINCOAT200@QQ.COM',bg='black',fg='#423900',font= ("Lucida Sans Typewriter", 8),height=2).pack()
root.mainloop()
os.remove("bg_gif")