from os import listdir
from chardet import detect
import os
import struct
import traceback
"""
修改文件内容
"""
def changeContent(path):
    with open(path,"rb+") as fp:#以二进制的形式打开文件
        bytes_content = fp.read()#获取文件二进制字节
        file_encoding = detect(bytes_content)["encoding"]#获取文件编码格式
        if fp.read() is None or len(bytes_content)==0:#判断文件是否为空
            print("文件为空")
            return
        file_content = bytes_content.decode(file_encoding)#将bytes解码成string
        # 修改文件内容
        file_content = file_content.replace("WEBVTT\n\n","").replace("<v.primary Chinese>","").replace("</v>","")\
            .replace("<v.secondary English>","").replace("<v English>","")
        fp.seek(0)#定位光标至文件开头
        fp.truncate()#清空当前光标所在位置之后的文件内容
        fp.write(file_content.encode(file_encoding))#将修改好后的内容写入文件
        print("已完成内容修改,文件编码格式：",file_encoding)
"""
修改文件编码
"""
def transformEncoding(path):
    with open(path,"rb+") as fp:
        target_encoding = "utf-8"#转换的目标编码格式
        content = fp.read()#获取文件二进制字节
        fileEncoding = detect(content)["encoding"]#获取字节的编码格式
        #windows-1252以及iso-8859-1部分特殊符号无法正常解码的问题，如：'
        if "Windows-1252" == fileEncoding:contentEncoding ="utf-8"
        elif "ISO-8859-1" == fileEncoding:contentEncoding = "gbk"
        else:contentEncoding = fileEncoding
        fp.seek(0)
        final_bytes = b""
        for line in fp.readlines():
            #处理目标编码格式无法编码的字符
            try:
                final_bytes += line.decode(contentEncoding).encode(target_encoding)#根据所需编码格式进行转码
            except Exception as e:
                print("该行转换编码失败，含有无法编码的特殊字符，内容："+line.decode(contentEncoding))
                final_bytes += line
        fp.seek(0)
        fp.truncate()
        fp.write(content)
        print("编码格式转换成功，将",fileEncoding,"转换成",target_encoding)
"""
根据拓展名查询某文件夹下所有的文件
返回：list（各个文件地址）
"""
def findFileByExtension(path):
    #递归查询
    def recursiveQuery(path):
        fns = (fn for fn in listdir(path))#获取该文件下所有文件
        for fn in fns:
            if judgeByExtension(fn,".srt"):
                file_path_list.append(os.path.join(path,fn))
            if os.path.isdir(os.path.join(path,fn)):
                recursiveQuery(os.path.join(path,fn))
    file_path_list = []  # 文件地址存储list
    recursiveQuery(path)
    return file_path_list
"""
根据文件名判断文件类型
"""
def judgeByExtension(file_name,extension):
    if file_name.endswith(extension):
        return True
    return False