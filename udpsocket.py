import socket
import threading
import time
from tkinter import INSERT

import openpyxl
import tkinter as tk
from tkinter.filedialog import askopenfilename

file: str
udp_flag: bool
flow_id: int
left_num: int


def get_now_time():
    now = time.localtime()
    now_time = time.strftime("%Y%m%d", now)
    return now_time


def get_device_id():
    print('filename is ', str(file))
    wb = openpyxl.load_workbook(str(file))
    sheet1 = wb['sheet1']
    nrows = sheet1.max_row + 1
    for j in range(1, nrows):
        value = sheet1.cell(j, 2).value
        if value == 0:
            global flow_id
            global left_num
            left_num = sheet1.max_row - j
            flow_id = j - 1
            print('剩余码数为：', left_num)
            # print('liushuihao ', flow_id, 'left number is', left_num)
            return sheet1.cell(j, 1).value, j


def set_token_used(k):
    wb = openpyxl.load_workbook(str(file))
    sheet1 = wb['sheet1']
    sheet1.cell(k, 2).value = "1"
    wb.save(str(file))
    pass


def start_udp():
    udp = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    print('start udp server...\n')

    print(socket.gethostbyname(socket.getfqdn(socket.gethostname())))
    try:
        udp.bind((socket.gethostbyname(socket.getfqdn(socket.gethostname())), 1777))
    except:
        pass
    global udp_flag
    udp_flag = True

    while udp_flag:
        rec_msg, addr = udp.recvfrom(1024)
        client_ip, client_port = addr
        print('client_ip:', client_ip, 'client_port:', client_port)

        print('msg from client:', rec_msg.decode('utf8'))
        if rec_msg.decode('utf8') == "recv":
            try:
                j = get_device_id()[1]
                set_token_used(j)
                print('烧录成功！\n')
            except:
                print('无可用序列号，请重新导入！！\n')
                pass

        if rec_msg.decode('utf8') == "ask1":
            try:
                ack_msg = get_device_id()[0]
                print(ack_msg)
                udp.sendto(ack_msg.encode('utf8'), addr)
            except:
                print('无可用序列号，请重新导入！！\n')
                pass

        if rec_msg.decode('utf8') == "ask2":
            ack_msg = get_now_time()
            print(ack_msg)
            udp.sendto(ack_msg.encode('utf8'), addr)

        if rec_msg.decode('utf8') == "ask3":
            try:
                ack_msg = get_device_id()[1] - 1
                ack_msg = str(ack_msg)
                print(ack_msg)
                udp.sendto(ack_msg.encode('utf8'), addr)
            except:
                print('无可用序列号，请重新导入！！\n')
                pass

        if not udp_flag:
            print('stop udp server\n')
            break


def stop_udp():
    global udp_flag
    udp_flag = False


def select_open_file():
    global file
    file = askopenfilename(title='选择文件', filetypes=[('excel文件', '*.xlsx')])
    print('file:', file)


def start_udp_thread():
    t = threading.Thread(target=start_udp, name='aa')
    t.start()


root = tk.Tk()
root.geometry("300x180")
root.title("ipc烧录工具")
file = './import.xlsx'
get_device_id()
tk.Button(root, text='选择文件', width=15, height=2, command=select_open_file).pack()
tk.Button(root, text="开始udpserver", width=15, height=2, command=start_udp_thread).pack()
tk.Button(root, text="停止udpserver", width=15, height=2, command=stop_udp).pack()

dstr = tk.StringVar()
lb = tk.Label(root, textvariable=dstr)
lb.pack()


def get_flow_id():
    global flow_id
    global left_num
    string = '流水号：', flow_id, '剩余码数：', left_num
    dstr.set(string)
    root.after(3000, get_flow_id)


get_flow_id()

root.mainloop()
