from stock import ChangeStrFile as cs

path = "D:\\BaiduNetdiskDownload\\Udacity深度学习\\01.神经网络"
for file_path in cs.findFileByExtension(path):
    print("开始转化，文件地址：", file_path)
    cs.changeContent(file_path)
    cs.transformEncoding(file_path)
    print("已完成转化，文件地址：",file_path)
    print("------------------------------------------------------------------------------")
