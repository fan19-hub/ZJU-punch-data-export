
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import datetime
from time import sleep
import json
import zipfile
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
# for the sys.exit() method call
import os,re
import pandas as pd
from re import compile,match
from urllib.request import URLopener, urlopen, urlretrieve
from bs4 import BeautifulSoup
# username='H620018'
# password='ly123456'
username=''
password=''
download_pth=os.getcwd()+'\\'+'疫情数据导出'

for i in range(5):
    if os.path.exists(download_pth):
        download_pth=os.getcwd()+'\\'+'疫情数据导出'+'(%i)'%i
    else:
        os.makedirs(download_pth)
        break
中文月份2num={"一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,"七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12}

def openPage(url):
      global download_pth
      """ 打开网页 """
      opt = Options()
      opt.add_argument('--disable-gpu')
      opt.add_experimental_option('prefs',{'download.default_directory':download_pth})      
      try:
            driver_path=r".\chromedriver.exe"
            browser = webdriver.Chrome(executable_path=driver_path, chrome_options = opt)
      except:
            print("\n错误：chromedriver出错，请检查：\n1、chromedriver是否与AutoHit在同一文件夹下 \n2、chromedriver版本是否与chrome浏览器版本匹配")
            get_chromedriver()
            return None
      try:
            if browser!=None:
                  browser.get(url)
      except:
            print("获取网页失败，可手动打开网页验证网络连接，并自行排除网络故障。提示：若使用过vpn，请在windows代理服务器设置里关闭代理即可）")
            browser.close()
            return None
      
      return browser

def login(browser)->bool:
    """ 浙大统一认证登录 """
    global username
    global password
    try:
        #输入用户名
        usr = browser.find_element_by_id("username")
        usr.send_keys(username)

        #输入密码
        psw = browser.find_element_by_id("password")
        psw.send_keys(password)

        #按提交按钮
        loginbtn = browser.find_element_by_id("dl")
        loginbtn.click()
    except:
        print("登录界面提交用户名、密码失败")
        assert -1
    #检查登录是否成功
    try:
        error =browser.find_element_by_xpath("//p[@id='errormsg']/span[@id='msg']")
        print("登录失败：",error.text)
        return 0
    
    except:
        return 1

def get_chromedriver()->bool:
        "自动下载合适版本的chromedriver"
        print("请按以下提示找到chrome浏览器的版本号并输入：\n\t1.打开chrome浏览器\n\t2.在输入网址的地方输入chrome://settings/help，回车")
        while 1:
                version=(input("(示例:91.0.4472.101)请输入您的版本号：")).strip(' ')
                if match(r"(\d{2}[.]\d[.]\d{4}[.]\d{2,3}|2[.]\d{1,2})",version)!=None:
                    break
                print("格式有误，请重新输入")
        link="http://npm.taobao.org/mirrors/chromedriver/"+version+"/chromedriver_win32.zip"
        ans='n'
        while('n' in ans or 'N' in ans):
            print("请确保您的网络连接正常")
            ans=input('您的网络连接是否正常？是(输入y并回车)否(输入n并回车)')
        try:
            urlopen(link)
        except:
            try:
                    _version=".".join(version.split(".")[:3])
                    version_pattern=compile('%s[.]\\d{3}'%_version)
                    html=urlopen("http://npm.taobao.org/mirrors/chromedriver")
                    bs = BeautifulSoup(html, 'html.parser')
                    newversion=(bs.find('a',string=version_pattern)).text[:-1]
                    print("镜像网站上没有该版本所对应的chromedriver，为您切换了相近版本：%s"%newversion)
                    link="http://npm.taobao.org/mirrors/chromedriver/"+newversion+"/chromedriver_win32.zip"
            except Exception as e:
                    print("报错！错误信息：",e)
                    return 0

        print("准备下载chromedriver:")
        try:
            urlretrieve(link,os.getcwd()+"/"+"chromedriver.zip",Schedule) 
            print("chromedriver压缩包下载成功！")
            fz = zipfile.ZipFile(os.getcwd()+"/"+"chromedriver.zip", 'r')
            for file in fz.namelist():
                    fz.extract(file,os.getcwd())
            print("\n\n获取成功！")
            return 1
        except:
            print("chromedriver下载或解压失败，请到网站自行查找，手动安装 http://npm.taobao.org/mirrors/chromedriver/")
            print("您的版本号是：%s"%version)
            return 0
         
def Schedule(a,b,c):
   '''
   a:已经下载的数据块
   b:数据块的大小
   c:远程文件的大小
   '''
   per = 100.0*a*b/c
   if per > 100:
      per = 100
   print('%.2f%%' % per)  

def get_calender():
    time_str_list=[]
    for i in range(30):
        time_str_list.append(str(datetime.date.today()-datetime.timedelta(days=i)))
    return time_str_list

def get_page(browser,time_str):
    try:
        url="https://healthreport.zju.edu.cn/site/epideZju/epitypeZju?type=total&date=%s"%time_str
        browser.get(url)
        js ='document.getElementsByClassName("upbtna")[0].click()' 
        browser.execute_script(js)
        js ='document.getElementsByClassName("btns")[0].children[1].click()'
        browser.execute_script(js)
        return 1
    except:
        print("Fail to process the page")
        return 0
    

def 一键导出():
    
    bs=openPage("https://healthreport.zju.edu.cn/site/epideZju/instruZju") 
    sleep(3)  
    a=login(bs)

    if bs==None or a!=1:      #if opening page or loging failed
        print("Fail to access")
        return 0
    else:
        time_str_list=get_calender()
        for time_str in time_str_list:
            a=get_page(bs,time_str)
            if a==0:
                return 0
            sleep(1)
    return 1
 
class application(Frame):
    '''application'''
    def __init__(self,master):
        super(application,self).__init__(master)
        
        self.grid()             #initialize the grid
        self.get_your_info()
        self.time_str_list=get_calender()
        self.dirname=download_pth
        self.create_widgets()   #add the widgets
    def get_your_info(self):
        """ 获取用户的用户名和密码（读取json文件。如果json文件中没有，就请求用户输入并存入json） """
        global username
        global password

        try:
            with open('userdata.json', 'r',encoding="utf-8") as f:
                    content = f.read()
            info_dict= json.loads(content)
            try:
                    username=info_dict["uname"]
                    password=info_dict["pwd"]
            except:
                    print("用户名或密码信息丢失。请修改")
                    info_dict=self.setting()
        except:
            info_dict=self.setting()
    def setting(self):
        screenheight=150
        screenwidth=300
        self.top = Toplevel()
        self.top.title('修改用户信息')
        self.top.geometry('%dx%d+%d+%d'%(300,100,(screenwidth-300)/2,(screenheight-150)/2))
        #user name
        Label(self.top,text='User name:').grid(row=1,column=0,sticky=E)                        
        self.e2=Entry(self.top,width=10)  #the input bar
        self.e2.grid(row=1,column=1,padx=1,pady=1) 
        #password
        Label(self.top,text='Password:').grid(row=2,column=0,sticky=E)                        
        self.e3=Entry(self.top,width=10)  #the input bar
        self.e3.grid(row=2,column=1,padx=1,pady=1) 
        Button(self.top, text='Submit',command=self.modify_json).grid(row=3,column=1,sticky=E)
    def modify_json(self):
        global username
        global password
        #请求用户输入：
        username=self.e2.get() 
        password=self.e3.get() 
        self.top.destroy()
        
        #创建一个用户信息的字典info_dict
        info_dict={}
        info_dict["uname"]=username
        info_dict["pwd"]=password

        #将信息存储到json中，之后自动调用
        with open('userdata.json', 'w',encoding='utf-8') as f:
                b=json.dumps(info_dict)
                f.write(b)
        return info_dict   
    def open(self):
        try:
            self.dirname=filedialog.askdirectory()
            if self.dirname=='':
                return
            messagebox.showinfo(title='报告',message= "打开成功！")

        except Exception as err:
            messagebox.showerror(title='Error',message= "ERROR:"+str(err))

    def process(self):
        dirlist=os.listdir(self.dirname)
        self.namelist=[name for name in dirlist if "疫情防控填报明细" in name]
        if len(self.namelist)==0:
            messagebox.showerror(title='错误',message="未发现疫情防控填报明细excel文件")
            return
        if len(self.namelist)<30:
            messagebox.showwarning(title='注意',message="疫情防控填报明细文件数量小于30，可能是还没下载完，可以等下载完后再重新生成一次")
        if len(self.namelist)>30:
            messagebox.showerror(title='错误',message="疫情防控填报明细文件数量大于30，可能是有之前下载的疫情防控明细文件，请先清理后重试")
            return
        li=["姓名"]+["最近上报时间%s"%self.time_str_list[n] for n in range(30)]+["所在地点%s"%self.time_str_list[n] for n in range(30)]
        df_res=pd.DataFrame(columns=li)
        for i in range(len(self.namelist)):
            
            df = pd.read_excel(self.dirname +'/'+self.namelist[i])
            # n is used to match the number in the filename like 疫情防控填报明细 (1).xlsx
            # Then we can use n to find the right column to write in
            print(self.namelist[i])
            n=re.search(r" \(([^\)]*)\)",self.namelist[i])
            print(n)
            if n==None:
                n=0
            else:
                n=int(n.group(1))
            df_res[["姓名"]]=df[['姓名']]
            df_res[["最近上报时间%s"%self.time_str_list[n]]]=df[['最近上报时间']]
            df_res[["所在地点%s"%self.time_str_list[n]]]=df[['所在地点']]
            print(self.namelist[i]+'********文件处理成功')
            print("最近上报时间%s"%self.time_str_list[n])
            # except Exception as e:
            #     messagebox.showerror(title='错误',message=self.namelist[i]+'********文件处理失败')
        try:
            df_res.to_excel(self.dirname +"/"+"0汇总表格.xlsx")
            print('写入完成')
            messagebox.showinfo(title='报告',message="处理完毕，请查看："+self.dirname +"/"+"0汇总表格.xlsx")
        except:
            messagebox.showerror(title='错误',message="写入"+self.dirname +"/"+"0汇总表格.xlsx"+"失败\n"+"如果您在excel中打开了此文件，请关闭后重试")  
            
    def submit(self):
        a=一键导出()
        if a==1:
            messagebox.showinfo(title='报告',message="导出成功！")
        else:
            messagebox.showinfo(title='报告',message="导出失败！")

    def help(self):
        txt="1.先点击一键导出，等待所有excel表格下载完成\n！！注意不要修改文件名！！\n2.在“开始”菜单栏点击“打开”，打开excel文件所在目录（建议创一个新文件夹，里面尽量只包含这次下载的文件）\n\n(c)2021 Fan Yang @ECE1902. All rights reserved. E-mail fan.19@intl.zju.edu.cn for help"
        messagebox.showinfo(title='帮助',message=txt)

    def quitt(self):
        q= messagebox.askquestion(title='Notice', message='Are you sure to quit?')
        if(q=='yes'):
            root.quit()

        
    def create_widgets(self):
        '''用来创建并定位我们所需的控件'''
        appmenu = Menu(self, tearoff=0) #1的话多了一个虚线，如果点击的话就会发现，这个菜单框可以独立出来显示
        appmenu.add_command(label="打开",command=self.open)
        appmenu.add_separator()
        appmenu.add_command(label="设置",command=self.setting)


        menubar = Menu(root)
        menubar.add_cascade(label="开始", menu=appmenu) #原理：先在主菜单中添加一个菜单，与之前创建的菜单进行绑定。
        menubar.add_command(label="使用说明",command=self.help)
        menubar.add_command(label="退出", command=self.quitt)
        root.config(menu=menubar)
 
        #Title
        line=0
        col=0 

        #Subtitle
        line+=1
        col=0
        padx=0
        self.label1=Label(self,text="     ",compound = CENTER,font=("times new roman",30),fg = "black")
        self.label1.grid(row=line,column=col ,sticky=W,ipadx=padx)   
        
        col+=1
        self.label2=Label(self,text="打卡数据导出",compound = CENTER,font=("times new roman",30),fg = "black")
        self.label2.grid(row=line,column=col ,sticky=E,ipadx=padx)   

        col+=1
        self.label3=Label(self,text="     ",compound = CENTER,font=("times new roman",30),fg = "black")
        self.label3.grid(row=line,column=col ,sticky=W,ipadx=padx)   

        #Q1
        line+=1
        col=0
        ipady=4
        self.button1=Button(self,text="一键导出",font=("times new roman",20),command=self.submit)
        self.button1.grid(row=line,column=col,sticky=E,ipady=ipady,pady=6) 

        col+=1

        col+=1
        self.button2=Button(self,text="生成汇总表",font=("times new roman",20),command=self.process)
        self.button2.grid(row=line,column=col,sticky=W,ipady=ipady,pady=6)
        # pass

        # col+=1
    #     self.input=Text(self,width=33,height=1,font=("times new roman",20),bg='Gainsboro')
    #     self.input.grid(row=line,column=col,sticky=W,columnspan=2)

    #     # col+=1
    #     # self.label=Label(self,text="Grading Helper",compound = CENTER,font=("times new roman",30),fg = "black")
    #     # self.label.grid(row=line,column=col ,sticky=W,columnspan=2)  


if __name__ == "__main__":  
    root=Tk()                   #创立一个根窗口对象，叫root
    root.title("打卡数据导出工具")       #set title, default: tk
    #root.geometry("400x500")
   
    '''background_image =  PhotoImage(file ="C:/Users/fan.19/Desktop/final project/OIP.png")
    w = background_image.width()
    h = background_image.height()
    root.geometry('%dx%d+0+0' % (w,h))
    
    background_label = Label(root, image=background_image)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)'''
    # 获取屏幕尺寸以计算布局参数，使窗口居屏幕中央
    width =537
    height =150
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(alignstr)
  
    app=application(root)       #create a application instance called app
    app.mainloop()              #run it    

        
        
    #    #Q2
    #     line+=1
    #     col=0
    #     ipady=4
    #     self.Q2=Label(self,text="Q2 ",compound = CENTER,font=("times new roman",20),fg = "black")
    #     self.Q2.grid(row=line,column=0 ,sticky=W) 


    #     col+=1
    #     self.n3=Entry(self,highlightcolor='red',relief=RAISED,width=25,font=("times new roman",20))  
    #     self.n3.grid(row=line,column=col,sticky=W,ipady=ipady,pady=0)


    #     #button
    #     line+=1
    #     col=0
    #     ipady=0

    #     self.button1=Button(self,text="Back",font=("times new roman",20),command=self.back)
    #     self.button1.grid(row=line,column=0,sticky=W,ipady=ipady,pady=6)
    #     self.button2=Button(self,text="Skip this",font=("times new roman",20),command=self.skip)
    #     self.button2.grid(row=line,column=2,sticky=W,ipady=ipady,pady=6)
    #     self.button1=Button(self,text="Save & Next",font=("times new roman",20),command=self.write)
    #     self.button1.grid(row=line,column=2,sticky=E,ipady=ipady,pady=6)

    #     # self.output=Text(self,width=25,height=2,bg='Gainsboro')
    #     # self.output.grid(row=5,column=1) 
    #     # self.output.insert("insert", "I love Python!")



# def get_calender():
#     time_tuple_list=[]
#     for i in range(30):
#         time_tuple_list.append(str(datetime.date.today()-datetime.timedelta(days=i)).split('-'))
#     return time_tuple_list
# def choose_date(browser,time_tuple):
#     year=int(time_tuple[0])
#     month=int(time_tuple[1])
#     day=int(time_tuple[2])
#     # wait for page opening
#     b=0
#     for i in range(11):
#         try:
#             js ='document.getElementsByClassName("alltitle")' 
#             browser.execute_script(js)  # 执行js语句
#             b=1
#             break
#         except:
#             sleep(1)
#         continue
#     if b==0:
#         print("用了太长时间打开网页，超时退出")

#     #click on the date
#     js ='document.getElementsByClassName("tip txt-input txt-buydate no-border-left text-right")[0].click()' 
#     browser.execute_script(js)  # 执行js语句

#     current_year=int(browser.find_element_by_xpath("//a[@class='md_headtext yeartag' and @href='javascript:void(0);']").text)
#     n=year-current_year
#     if n<0:
#         for i in range(-n):
#             js ='document.getElementsByClassName("md_prev change_year")[0].click()' 
#             browser.execute_script(js)  # 执行js语句
#     if n>0:
#         for i in range(n):
#             js ='document.getElementsByClassName("md_next change_year")[0].click()' 
#             browser.execute_script(js)  # 执行js语句
        

#     current_month=中文月份2num[browser.find_element_by_xpath("//a[@class='md_headtext monthtag' and @href='javascript:void(0);']").text]
#     n=month-current_month
#     if n<0:
#         for i in range(-n):
#             js ='document.getElementsByClassName("md_prev change_month")[0].click()' 
#             browser.execute_script(js)  # 执行js语句
#     if n>0:
#         for i in range(n):
#             js ='document.getElementsByClassName("md_next change_month")[0].click()' 
#             browser.execute_script(js)  # 执行js语句
    
#     # js ='document.getElementsByName("area")[0].children[1].click()' 
#     try:
#         a=browser.find_element_by_xpath("//li[@data-day='%i']"%day)
#         a.click()
#     except:
#         print("ddd")
#     js ='document.getElementsByClassName("md_next change_month")[0].click()' 
#     browser.execute_script(js)  # 执行js语句
#     #click on the date to quit
    
#     js ='document.getElementsByClassName("tip txt-input txt-buydate no-border-left text-right")[0].click()' 
#     browser.execute_script(js)  # 执行js语句 