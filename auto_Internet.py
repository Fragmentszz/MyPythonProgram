
import requests
import os
import re
from pywifi.wifi import PyWiFi
from pywifi.profile import Profile
from time import sleep

sleep(1)

def cs_config(libraries_path):
    file_name = libraries_path + '/配置.txt'
    with open(file_name,'r',encoding='utf-8') as fp:
            xh=fp.readline().strip('\n')[3:]
            mm=fp.readline()[3:]
    if xh !='' and mm !='':
        return xh,mm
    else:
        return 0

def initial(libraries_path ):
    # 访问libraries文件夹
    try:

        file_name = libraries_path+'/配置.txt'
        with open(file_name,'w',encoding='utf-8') as fp:
            xh = input("请输入账号")
            mm = input("请输入密码")
            str = "账号:"+xh+"\n"+"学号:"+mm
            fp.write(str)
        return (xh,mm)
    except Exception as e:
        print(e)
        return 0





def cs_wlan():
    wifi=PyWiFi()
    iface = wifi.interfaces()[0]

    return iface.status()




def connect():
    wifi=PyWiFi()
    iface = wifi.interfaces()[0]
    profile = Profile()
    for z in iface.network_profiles():
        if z.ssid == 'ahu.portal' :
            profile = z
            break
    iface.connect(profile)
    sleep(1)







if cs_wlan() == 0:
    connect()


libraries_path = '.'


if cs_config(libraries_path) == 0:
    user_account , user_password = initial(libraries_path)
else :
    user_account , user_password  = cs_config(libraries_path)



text=os.popen('ipconfig/all').read()
list=text.split()
while True:
    try:
        list.remove('.')
    except:
        break
print(list)

wlan_index=list.index('WLAN:')
url='http://172.16.253.3:801/eportal/?'
wlan_user_mac=''.join(list[wlan_index+13].split('-'))
wlan_user_ip=list[wlan_index+29][:-4]

with open('cs.txt','w') as fp:
    fp.write(wlan_user_ip)
    fp.write(wlan_user_mac)
    fp.close()


param={
    'c': 'Portal',
    'a': 'login',
    'callback': 'dr1003',
    'login_method': '8',
    'user_account': user_account,
    'user_password': user_password,
    'wlan_user_ip': wlan_user_ip,
    'wlan_user_ipv6': '',
    'wlan_user_mac': wlan_user_mac,
    'wlan_ac_ip': '172.20.0.165',
    'wlan_ac_name':'',
    'jsVersion': '3.3.2',
    'v': '6329'
}



header={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.55'
}


text = requests.get(url=url,params=param,headers=header).text


pattern = '"ret=\'(.*?)\';"'
pattern2 = '"result":"(.*?)"'
result = re.findall(pattern=pattern,string=text)[0]
result2 = re.findall(pattern=pattern2,string=text)[0]

if result2 == '1':
    print("成功了哦~")
elif  'userid' in result:
    print("账号或者是密码错误咯")
elif result == 'no errcode':
    print("已经登录咯")

a=input("输入")




