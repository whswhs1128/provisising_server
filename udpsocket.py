import socket
import time
from excel_format import get_device_id, set_token_used


def get_now_time():
    now = time.localtime()
    now_time = time.strftime("%Y%m%d", now)
    return now_time


udp = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
print('start udp server...\n')

udp.bind(('192.168.1.104', 1777))

while True:
    rec_msg, addr = udp.recvfrom(1024)
    client_ip, client_port = addr
    print('client_ip:', client_ip, 'client_port:', client_port)

    print('msg from client:', rec_msg.decode('utf8'))
    if rec_msg.decode('utf8') == "recv ok":
        try:
            j = get_device_id()[1]
            set_token_used(j)
        except:
            print('无可用序列号，请重新导入！！\n')
            pass

    if rec_msg.decode('utf8') == "ask id":
        try:
            ack_msg = get_device_id()[0]
            udp.sendto(ack_msg.encode('utf8'), addr)
        except:
            print('无可用序列号，请重新导入！！\n')
            pass

    if rec_msg.decode('utf8') == "ask date":
        ack_msg = get_now_time()
        udp.sendto(ack_msg.encode('utf8'), addr)

    if rec_msg.decode('utf8') == "ask flow id":
        try:
            ack_msg = get_device_id()[1]
            ack_msg = str(ack_msg)
            udp.sendto(ack_msg.encode('utf8'), addr)
        except:
            print('无可用序列号，请重新导入！！\n')
            pass
