import imaplib
import email
import base64
import datetime
import time
from tkinter import Tk,Label,messagebox
import win32com.client
import winsound
import threading
import os

def getMailContent(host,port,user,password,date_start,str_subject,path_resultHtml):
    # 连接到服务器
    conn=imaplib.IMAP4_SSL(host=host,port=port)     #993加密连接；143未加密

    conn.login(user, password)

    # print(conn.list())

    # 选择发件箱
    conn.select('(xxxxxx)')

    # 筛选今日以后的邮件
    _,date_mailid=conn.search(None,'SINCE %s' % date_start.strftime('%d-%b-%Y'))     #1-JAN-2023
    # print(data[0])

    list_mailid=date_mailid[0].split()
    # print(list_mailid)

    str_content=''

    for i in range(len(list_mailid)-1,-1,-1):
        # print(i)
        mailid=list_mailid[i]

        # 获取邮件内容
        _,maildata=conn.fetch(mailid, '(BODY[])')   # RFC822等同BODY[]

        # 转换成message类型
        msg=email.message_from_bytes(maildata[0][1])
        # msg=email.message_from_string(maildata[0][1].decode('utf-8'))
        # print(msg)

        # 获取主题
        subject_decode=email.header.decode_header(msg.get('subject'))[0][0]

        if isinstance(subject_decode, str):
            subject=subject_decode
        else:
            subject=subject_decode.decode('utf-8')

        # print(subject)

        if subject==str_subject:
            # 获取邮件正文
            if msg.is_multipart():
                
                for payload in msg.get_payload():
                    # print(type(payload))
                    str_content +=str(payload)+'\n\n'
                    # print(str_content)
            else:
                str_content=msg.get_payload(decode=True)

            tag_bodyStart='Content-Transfer-Encoding: base64'
            index=str_content.find(tag_bodyStart)
            if index !=-1:
                str_content=str_content[index+len(tag_bodyStart):].strip()

                # base64解密
                str_content=base64.b64decode(str_content).decode('utf-8')

            # print(str_content)
            with open(path_resultHtml,'w',encoding='utf-8') as f:
                f.write(str_content)

            break
    return str_content



def alert(str_tips,datetime_End):
    global flag_run
    flag_run=True

    # 窗口提示界面，tkinter会阻塞线程，所以开新线程
    threading.Thread(target=windowTips,args=(str_tips,str_tips)).start()

    # 蜂鸣声和下面的Windows发音并行运行，所以开新进程
    threading.Thread(target=soundTips,args=(datetime_End,)).start()

    # 调用Windows发音
    speak = win32com.client.Dispatch('SAPI.SPVOICE')
    while datetime.datetime.now() < datetime_End:
        if flag_run==False:break
        # print(datetime.datetime.now())
        speak.Speak(str_tips)
        time.sleep(1)


def closeWindow(mainWindow):
    if messagebox.askyesno('是否关闭提示？','是否关闭提示？'):
        global flag_run
        flag_run=False
        mainWindow.destroy()
        # mainWindow.quit()

def windowTips(str_title,str_tips):
    mainWindow = Tk()
    # 主窗口设置
    mainWindow.title(str_title)
    mainWindow.attributes('-topmost',True)  # 窗口置顶

    mainWindow.protocol('WM_DELETE_WINDOW',lambda:closeWindow(mainWindow))  # 绑定窗口关闭事件

    bgColor= '#42586e' # '#9fefe2' # '#94E5C0' # '#7CFC00'
    mainWindow.config(background=bgColor)

    width_win = 300
    height_win = 230

    x_win = (mainWindow.winfo_screenwidth() // 2) - (width_win // 2)
    y_win = (mainWindow.winfo_screenheight() // 3) - (height_win // 2)

    mainWindow.geometry('{}x{}+{}+{}'.format(width_win, height_win, x_win, y_win))  # 窗口居中，设置 窗口大小、位置：字符串格式：width x height + x + y

    mainWindow.resizable(False,False)   # 禁止修改窗口大小

    lbl=Label(master=mainWindow,text=str_tips,font=('黑体',24),foreground='#e3d4af',background= bgColor)
    lbl.pack(expand=True)

    # 显示窗口
    mainWindow.mainloop()


def soundTips(datetime_End):
    while datetime.datetime.now() < datetime_End:
        global flag_run
        if flag_run==False:break
        # print(datetime.datetime.now())
        # 调用windowsBeep

        # 生日快乐歌-->
        # winsound.Beep(523, 200)
        # winsound.Beep(523, 200)
        # winsound.Beep(578, 400)
        # winsound.Beep(523, 400)
        # winsound.Beep(698, 400)
        # winsound.Beep(659, 800)
 
        # winsound.Beep(523, 200)
        # winsound.Beep(523, 200)
        # winsound.Beep(578, 400)
        # winsound.Beep(523, 400)
        # winsound.Beep(784, 400)
        # winsound.Beep(698, 800)
 
        # winsound.Beep(523, 200)
        # winsound.Beep(523, 200)
        # winsound.Beep(1046, 400)
        # winsound.Beep(880, 400)
        # winsound.Beep(698, 400)
        # winsound.Beep(659, 400)
        # winsound.Beep(578, 400)
 
        # winsound.Beep(932, 200)
        # winsound.Beep(932, 200)
        # winsound.Beep(880, 400)
        # winsound.Beep(698, 400)
        # winsound.Beep(784, 400)
        # winsound.Beep(698, 800)
        # <--生日快乐歌

        # 周杰伦 夜曲-->
        time.sleep(0.5)            #0
        winsound.Beep(880, 250)     #6
        winsound.Beep(987,250)      #7

        winsound.Beep(1046,500)     #高1
        winsound.Beep(1046,250)     #高1
        winsound.Beep(1046,250)     #高1
        winsound.Beep(1046,250)     #高1
        winsound.Beep(1046,750)    #高1

        winsound.Beep(987,500)      #7
        winsound.Beep(1318,250)     #高3
        winsound.Beep(1318,250)     #高3
        winsound.Beep(1318,1000)    #高3

        winsound.Beep(1760,500)     #高6
        winsound.Beep(1760,250)     #高6
        winsound.Beep(1760,250)     #高6
        time.sleep(0.5)             #0
        winsound.Beep(1567, 500)    #高5
        winsound.Beep(1396, 250)    #高4
        winsound.Beep(1567, 750)    #高5
        winsound.Beep(1046,250)     #高1
        winsound.Beep(1046,1000)    #高1
        # <--周杰伦 夜曲

        # print('end')
        time.sleep(1)


def main(host,port,user,password,date_start,str_subject,path_resultHtml):
    # 获取邮件内容
    str_content=getMailContent(host,port,user,password,date_start,str_subject,path_resultHtml)

    # 告警超时时间
    str_date=str(datetime.date.today())
    str_stopTime='7:40'
    str_dateFormat='%Y-%m-%d %H:%M'
    datetime_End=datetime.datetime.strptime(str_date+' '+str_stopTime,str_dateFormat)
    # print(datetime_End)

    if str_content=='':
        str_tips='警告!无邮件！'
        alert(str_tips,datetime_End)
        return

    with open(path_resultHtml,'r',encoding='utf-8') as f:
        str_content=f.read()

    # 打开结果文件
    os.startfile(path_resultHtml)

    str_child='成功执行：'
    index_start=str_content.find(str_child)
    if index_start !=-1:
        # 找到'成功执行:'时
        index_start+=len(str_child)
        index_end=str_content.find('；')
        str_num_success=str_content[index_start:index_end]

        count_rows=str_content.count('</tr>')-1
        print(count_rows,str_num_success)

        # 成功执行小于数据行
        if int(str_num_success)<count_rows:
            # 告警
            str_tips='警告!xxx流程失败！'
            alert(str_tips,datetime_End)
        else:
            if '无数据' in str_content:
                print('无数据' in str_content)
                str_tips='警告!xxx表无数据！'
                alert(str_tips,datetime_End)
    else:
        str_tips='错误!邮件内容错误！'
        alert(str_tips,datetime_End)



if __name__=='__main__':
    host='imap.xxx.com'
    port=143
    user='xxxxx@xxx.com'
    password='xxxxxxxxxxxx'
    date_start=datetime.date.today()
    str_subject='ssssssss'
    path_resultHtml='result_xxx.html'

    main(host,port,user,password,date_start,str_subject,path_resultHtml)
