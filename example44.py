# 批量从上市公司年报中获取指定内容已完成

# os模块： 用于处理文件路径，例如检查文件是否存在、创建目录等操作。
import os
print("hello,moon")
# (定义文件夹路径)这里设置了变量’path‘,他存储了文件夹的名称’年报‘，在windows操作系统中，文件夹路径可以用反斜杠’\‘来分割目录。
path='年报'  #文件所在文件夹

# (获取文件列表)os.listdir()函数返回指定路径(这里是年报文件夹)中所有文件和文件夹的列表。
# (拼接文件路径)使用了列表推导式来遍历文件夹中的每个文件和文件夹的名称'i'，然后将他们与文件夹路径path拼接起来，
# 在windows上，文件夹路径和文件名之间要使用反斜杠’‘
files = [path+"\\"+i for i in os.listdir(path)] #获取文件夹下的文件名,并拼接完整路径的列表files。

# pdfplumber模块： 用于处理PDF文件，提取文本和表格数据。
import pdfplumber

# time模块是python标准库中的一个模块，主要用于处理与时间相关的功能。
import time

# 1.获取当前时间戳，使用time.time()函数可以获取当前时间戳。
time0= time.time()

#从字符串中提取指定首尾的文字
# 定义了一个函数Gtt_text，用于从给定的source_str中提取位于start_str和end_str之间的字符串。
# 1.start_str:要搜索的起始关键词字符串。
# 2.end_str:要搜索的结束关键词字符串。
# 3.source_str:源字符串，从里面提取目标文本
def Get_text(start_str, end_str, source_str):
    # (找到起始位置)使用find()方法在source_str中查找start_str的位置索引。
    start = source_str.find(start_str) #找到开始关键词对应的位置索引
    # (截取目标文本)
    # 如果找到了start_str，则将start索引向后移动start_str的长度，则便从其后开始搜索end_str，
    # 使用find()方法在source_str中查找end_str的位置索引，
    # 如果找到了end_str，则通过切片操作source_str[start:emd]截取start_str和end_str之间的子字符串，
    # 使用strip()方法去除截取结果的首尾空白字符(空格，制表符等)，
    # (返回结果)如果成功找到了 start_str 和 end_str，则返回截取的子字符串；否则返回 None（因为没有显式的 else 分支）。
    if start >= 0:
        start += len(start_str)
        end = source_str.find(end_str, start)#找到结束关键词对应的位置索引
        if end >= 0:
            return source_str[start:end].strip() #截取起始位置之间的字符

#定义写入txt的函数
# （函数定义和参数说明）这段代码定义了一个函数 To_txt，
# 其目的是将给定的文本数据列表 final_text 写入到指定路径的文本文件中。让我们逐步解释每个部分的功能和每行代码的作用。
# 参数解释：
# filename: 要写入的文件路径，不包括扩展名 .txt，例如 'output'。
# final_text: 要写入文件的文本数据列表，每个元素代表一行文本。
def To_txt(filename, final_text):#filename为写入文件的路径，data为要写入数据列表.
    # （打开文件）使用 open() 函数打开文件，以写入模式 'w' 打开，并指定文件编码为 UTF-8。如果文件已存在，
    # 会清空文件内容；如果文件不存在，会创建新文件。
    file = open(filename + '.txt','w',encoding="utf-8")
    # （写入文件名）将filename写入文件的第一行，并在末尾添加换行符’\n‘
    file.write(filename + "\n")
    # （逐行写入文本数据）使用 for 循环遍历 final_text 列表中的每个文本行，将其写入文件中。如果不是列表中的最后一行文本，则在每行的末尾添加换行符 \n。
    for i in range(len(final_text)):
        text = final_text[i]
        if i != len(final_text)-1: #判断是否最后一个元素
            text = text+'\n'   #若不是最后一个元素才换行
        file.write(text)
    # （延时操作）加入一个短暂的延时（0.1秒），这是为了避免在批量写入大量文件时可能出现的文件乱码问题。这种情况在一些系统或环境中可能会发生。
    time.sleep(0.1) #加入一个延时，避免批量写入出现乱码
    # （关闭文件）使用 file.close() 方法关闭文件，确保所有写入操作完成并释放文件资源。
    file.close()
    
#获取年报中的“主要业务”信息
# (遍历文件列表)这里files是一个包含文件路径的列表，每个文件都是pdf文件。
for file in files:
    # （打开pdf文件并提取文本）
    # 使用 pdfplumber 库打开 PDF 文件，并迭代处理页码从 6 到 25 的页面（索引从 0 开始）。
    # 使用 page.extract_text() 提取每一页的文本内容，并将其存储在 data 列表中。
    # 如果在任何一页中找到了 key_words（例如 "重大变化情况"），则停止提取，以节省时间和资源。
    data = []
    key_words = "重大变化情况"
    with pdfplumber.open(file) as p:
        for i in range(6,26): #公司主要业务主要年报的在8~23页范围内
            page = p.pages[i] #选页
            page_text = page.extract_text() #提取文字
            data.append(page_text) #将提取的文字加入列表
            if key_words in page_text: #到结束关键词即结束抓取信息，避免浪费时间
                break # 终止for循环        

    # (合并数据列表为大字符串)将列表data中的所有字符串元素连接成一个大字符串source_str，这个大字符串包含了从多个文档提取的所有文本内容。
    #将数据列表`data`转换成一个大字符串
    source_str = "".join(data)
    #截取文字
    # (截取所需文本段落)调用 Get_text 函数，根据 start_str 和 end_str 截取 source_str 中介于这两者之间的文本段落。
    start_str = "公司业务概要"
    end_str = "重大变化情况"
    text_wanted = Get_text(start_str, end_str, source_str)
    #去掉不需要的尾巴
    # （处理最终文本）将 text_wanted 按照换行符 \n 分割成行，并移除最后一个空行（如果存在）。这是通过 [:-1] 来实现的，它表示从列表的第一个元素到倒数第二个元素。
    final_text = text_wanted.split("\n")[:-1]
    # （构建新文件路径）这里假设 file 是一个文件路径，通过 split("\\")[1] 获取文件名部分，并移除文件扩展名（假设文件扩展名为 .pdf）。
    new_file = "主要业务\\" + file.split("\\")[1][:-4]
    # （调用写入函数）调用 To_txt 函数，将处理后的 final_text 写入到以 new_file 命名的文本文件中。
    To_txt(new_file,final_text)
    # （输出处理完成信息）打印出处理完成的提示信息，显示已处理的文件名。
    print("{} 处理完成！".format(new_file))
    
time1= time.time()
# （计算时间差）计算开始时间time0和结束时间time1之间的时间差，并将其格式化输出，这个时间差表示程序的总执行时间。
print("处理完成，共用时 {} 秒。".format(time1-time0))