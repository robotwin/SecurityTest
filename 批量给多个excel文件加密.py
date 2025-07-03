# 'os'模块提供与操作系统交互的功能，例如文件路径操作等。
# 'sys'模块提供了对python运行时环境的访问和控制。
import os, sys

# 'win32com.client'这是一个第三方模块，用于windows平台上与COM组件进行交互，
# 例如，它可以用来操作Microsoft Office应用程序，如excel,word等，通常需要在系统中，
# 安装相应的Microsoft Office 以及安装 pywin32 库来使用。
import win32com.client

# 'time'这是一个python的标准模块，提供了与时间相关的函数，例如延时等待。
import time

# 'hashlib'这是python的标准模块，用于加密散列算法的接口
import hashlib

# ’tkinter‘这是python的标准库模块，用于创建图形界面(GUL)应用程序
# tk 是 tkinter 的别名，用于引入 tkinter 模块。

# from tkinter import filedialog 引入了 tkinter 模块中的 filedialog 子模块，
# 用于创建文件对话框，让用户选择文件或保存文件,并且起别名为tk
import tkinter as tk  # 调用GUI图形模块

# 从tkinter中导入filedalog模块时，可以使用他来创建文件操作的对话框，这些对话框允许用户交互的选择文件或目录
from tkinter import filedialog

# 作用:这个方法的作用是对指定的Excel文件进行密码保护,并将其另存为新文件.
# 参数解释
# 1.'old_filename':原始Excel文件的路径,即要进行密码保护的文件.
# 2.'new_filename':另存为新的Excel文件的路径,即保存密码保护后的文件
# 3.'pwd_str':要设置的保存时的访问密码.
# 4.'pw_str':可选参数,表示打开文件的密码,即为空字符串
def pwd_xlsx(old_filename, new_filename, pwd_str, pw_str=''):

    # 这一行通过com(Component Object Model)初始化了Excel应用程序，使得Python可以与Excel进行交互
    xcl = win32com.client.Dispatch("Excel.Application")

    # 打开指定名称为‘old_filename’的Excel工作簿。
    # ‘False,False’表示不更新链接，不以只读模式打开。
    # ‘pw_str’是可选的打开密码，用于打开密码保护的工作簿。
    wb = xcl.Workbooks.Open(old_filename, False, False, None, pw_str)

    # 禁止Excel弹出的警告和确认对话框，以便程序在后台运行时不会被这些弹窗打断。
    xcl.DisplayAlerts = False

    # 将已打开的工作簿('wb')，另存为新文件名'new_filename'。
    # 'pwd_str'用作设置工作簿访问权限的密码。
    # ’‘表示不设置写保护密码。
    # 保存时可设置访问密码.
    wb.SaveAs(new_filename, None, pwd_str, '')

    # 这一行实际上是一个没有实际效果的休眠操作，可能原本是为了引入一些延迟，但是休眠时间为0秒，不会产生实际作用。
    time.sleep(0)

    # 分别关闭工作簿并退出Excel应用程序,释放资源.
    wb.Close()
    xcl.Quit()

# 作用:定义了一个名为'read_path'的函数,其作用是读取指定路径'path'下的所有文件和文件夹名称,并将他们列表形式返回.
# 定义了一个函数'read_path',他有一个参数'path',表示要读取的目录路径.
def read_path(path):

    # 使用'os.listdir()'函数列出指定路径'path'下所有文件和文件夹名称,并将结果赋值给变量'dirs'
    dirs = os.listdir(path)
    # 将包含文件和文件夹名称的列表'dirs'返回给调用者
    return dirs

# 作用:这个函数的作用是扫描指定文件夹中的所有文件,筛选出后缀为'.xlsx'或'.xls'的Excel文件,并返回这些文件的完整路径列表.
# '方法_获取文件路径'是一个函数定义,他接受一个参数'data_dir',表示要搜索大的文件夹路径.
def 方法_获取文件路径(data_dir):

    # 'os.listdir(data_dir)'返回指定目录'data'中所有文件和文件夹的名称列表,并将结果赋值给'文件集'变量
    文件集 = os.listdir(data_dir)

    # 'filename'是一个空列表,用于存储条件的excel文件路径.
    filename = []

    # 使用'for'循环遍历'文件集'中的每个文件名,将当前文件名按照'.'分割成列表'文件名分解'
    for 文件名 in 文件集:
        文件名分解 = 文件名.split('.')

        # 检查'文件名分解'列表的最后一个元素(即文件扩展名),如果是'xlsx'或'xls',则将该文件的完整路径('data_dir+文件名')添加到'dilename'列表中.
        if 文件名分解[-1] == 'xlsx' or 文件名分解[-1] == 'xls':
            excel文件路径 = (data_dir + 文件名)
            filename.append(excel文件路径)
            # print(excel文件路径)

    # 返回存储了所有符合条件的Excel文件路径的'filename'列表.
    # print(filename)
    return filename

# def main是一个常见的python函数定义，通常用于标识程序的入口点
def main():

    # 创建了一个名为root的tkinter应用程序窗口对象
    root = tk.Tk()

    # withdraw()方法将窗口隐藏起来，这样用户就看不到他，但仍然可以使用他
    root.withdraw()

    # 这行代码打开一个对话框，让用户选择一个文件夹(通过askdirectory方法)，并将用户选择的文件夹路径存储在data_dir变量中，
    # title参数设置对话框的标题，最后+‘/’是为了确保data_dir以斜杠结尾，以便后续操作处理文件路径时的一致性
    data_dir = filedialog.askdirectory(title='请选择excel所在文件夹') + '/' #文件夹里面不要有除了excel外的任何文件

    # 接下来的步骤调用一个名为方法_获取文件路径的函数,(假设是自定义的函数),并将data_dir作为参数传递给这个函数,
    # 以获取该文件夹中所有文件的路径列表，并将结果存储在file_list变量中。
    # 将源文件路径里面的文件转换成列表file_list
    file_list = 方法_获取文件路径(data_dir)

    # 这里定义了一个名为dirs的字符串变量，其值为临时文件。
    dirs = '临时文件'

    # 然后代码使用os.path.exists()函数检查当前目录下是否存在名为‘临时文件’的文件夹
    if not os.path.exists(dirs):
        # 如果不存在，使用os.makedirs()函数创建这个文件夹，注意os.makdirs(),会创建所有必须的中间目录，以确保这个路径都存在。
        os.makedirs(dirs)  # 输出文件夹

    # 将‘dirs’变量与斜杠‘/’连接起来，将结果赋值给‘result_dir’变量，这样做确保文件夹路径木为有斜杠，以便后续操作路径拼接的一致性。
    # global result_dir
    result_dir = dirs + '/'

    # 将之前选择的源文件路径‘data_dir’存储在‘source’变量中，供后续使用
    # 源文件路径
    source = data_dir

    # 这行代码打开一个文件夹选择对话框，让用户选择一个木匾文件夹，并将选择的路径存储在ob变量中，'title'参数设置对话框的变态，
    # 同样的,末尾的'+''/'确保路径末尾有斜杠,以保持路径的一致性.
    # 目标文件路径
    ob = filedialog.askdirectory(title='请选择输出另存的文件夹') + '/'

    a = 0  # 列表索引csv文件名称放进j_list列表中，索引0即为第一个csv文件名称
    j_list = read_path(source)  # 文件夹中所有的csv文件名称提取出来按顺序放进j_list列表中
    # print("---->", read_path(source))       # read_path(source) 本身就是列表
    # print("read_path(source)类型：", type(read_path(source)))
    # 建立循环对于每个文件调用excel_to_csv()

    # 这是一个循环,遍历'file_list'中的每个原色'it',每个元素应该是一个文件的路径或名称.
    for it in file_list:
        # 打印当前文件开始加密的提示信息.
        print(it + '开始加密')
        # 根据索引'a'从'j_list'中获取当前文件对应的新文件名(假设'j_list'是一个包含新文件名的列表).
        j = j_list[a]  # 按照索引逐条将csv文件名称赋值给变量j
        # print(j)
        # 给目标文件新建一些名字列表
        # 构造新文件名的完整路径,'ob'是一个文件路径的基础目录.
        new_filename = ob + '\\' + str(j)
        # 设置新的加密密码.
        pwd_str = '654321'  # 新密码自定义，需修改
        # 处理文件路径中的斜杠,将路径中的斜杠'/',替换为反斜杠'\'
        aa = it
        # print(aa.replace("/", "\\"))
        bb = ob + str(j)
        # print(bb.replace("/", "\\"))
        # 替换斜杠后的文件路径,得到'path1'和'path2'.
        path1 = aa.replace("/", "\\")
        path2 = bb.replace("/", "\\")
        # Remove_password_xlsx(aa.replace("/", "\\"), pw_str)
        # 'try'中调用pwd_xlsx(path1,path2,path3)函数,尝试使用给指定的路径加密Excel文件,
        # 如果加密过程中出现异常没捕获异常并打印加密失败的提示信息,并继续处理下一个文件.
        try:
            pwd_xlsx(path1, path2, pwd_str)
            # 在加密成功后打印加密完成的提示信息
            print(it + '加密完成')
            # 程序休眠1秒,是为了控制处理速度或等待Excel操作完成.
            time.sleep(1)

        except Exception as e:
            print(it + '加密失败，请手动设置')

            pass
        # 更新索引'a',以便在下次循环中'j_list'中获取下一个文件的新名称
        a = a + 1

# if __name__ == '__main__':是python中一个常见的约定,用来判断是否直接执行当前脚本文件(作为主程序),
# 而不是作为模块被导入到其他脚本中,这个条件通常用来确保某些代码直接执行时运行,而在被导入是不执行.
if __name__ == '__main__':
    main()

#运行时第一个对话框是原始文件所在的位置，第二次对话框选择输出文件的位置；