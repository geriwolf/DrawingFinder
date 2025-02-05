import base64
# ico转base64
open_icon = open("./img/logo.ico", "rb")
b64str = base64.b64encode(open_icon.read())  # 转换为base64编码
open_icon.close()
write_data = "imgBase64 = %s" % b64str
f = open("./img/logo.py", "w+")
f.write(write_data)# 写入文件
f.close()