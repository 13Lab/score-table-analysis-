import xlrd
from matplotlib import pyplot as plt
ax = plt.gca()

L = []
kinds = []
x = []
def read_xlsx(c0,c1,c2,c3,c4,c5,c6,c7,xls_path,sheet_name,name_line,nums_of_kinds,width,n1,n2,n3,n4,n5,n6,n7,times):
    plt.cla()
    plt.clf()
    plt.close()
    draw_pic = times
    for draw_pic_num in range(3,draw_pic):
        data = xlrd.open_workbook(xls_path)
        data.sheet_names()
        table = data.sheet_by_name(sheet_name)
        L = table.row_values(draw_pic_num)
        # num_cols = table.ncols
        for i in range(1,name_line):
            del L[0]
        N_long = len(L)
        for i in range(0,N_long):
            c0.append(L[i])
        C_long = len(c0)
        if(C_long % (width * nums_of_kinds) == 0):
            if (nums_of_kinds >= 1):
                for i in range(0,width):
                    c1.append(c0[i])
            if (nums_of_kinds >= 2):
                for i in range(width, 2*width):
                    c2.append(c0[i])
            if (nums_of_kinds >= 3):
                for i in range(2*width, 3*width):
                    c3.append(c0[i])
            if (nums_of_kinds >= 4):
                for i in range(3*width, 4*width):
                    c4.append(c0[i])
            if (nums_of_kinds >= 5):
                for i in range(4*width, 5*width):
                    c5.append(c0[i])
            if (nums_of_kinds >= 6):
                for i in range(5*width, 6*width):
                    c6.append(c0[i])
            if (nums_of_kinds >= 7):
                for i in range(6*width, 7*width):
                    c7.append(c0[i])
        else:
            print("数据异常或不符合要求1")

        for i in range(1,width+1):
            x.append(i)
        true_num = width * nums_of_kinds
        if (C_long % true_num == 0):
            # for i in range(0,true_num):
            if(nums_of_kinds == 1):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
            elif(nums_of_kinds == 2):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
            elif(nums_of_kinds == 3):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
                plt.plot(x, c3, lw=1, c='orange', marker='s', ms=4, label=n3)
            elif (nums_of_kinds == 4):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
                plt.plot(x, c3, lw=1, c='orange', marker='s', ms=4, label=n3)
                plt.plot(x, c4, lw=1, c='aqua', marker='s', ms=4, label=n4)
            elif (nums_of_kinds == 5):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
                plt.plot(x, c3, lw=1, c='orange', marker='s', ms=4, label=n3)
                plt.plot(x, c4, lw=1, c='aqua', marker='s', ms=4, label=n4)
                plt.plot(x, c5, lw=1, c='m', marker='s', ms=4, label=n5)
            elif (nums_of_kinds == 6):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
                plt.plot(x, c3, lw=1, c='orange', marker='s', ms=4, label=n3)
                plt.plot(x, c4, lw=1, c='aqua', marker='s', ms=4, label=n4)
                plt.plot(x, c5, lw=1, c='m', marker='s', ms=4, label=n5)
                plt.plot(x, c6, lw=1, c='darkgreen', marker='s', ms=4, label=n6)
            elif (nums_of_kinds == 7):
                plt.plot(x, c1, lw=1, c='red', marker='s', ms=4, label=n1)
                plt.plot(x, c2, lw=1, c='blue', marker='s', ms=4, label=n2)
                plt.plot(x, c3, lw=1, c='orange', marker='s', ms=4, label=n3)
                plt.plot(x, c4, lw=1, c='aqua', marker='s', ms=4, label=n4)
                plt.plot(x, c5, lw=1, c='m', marker='s', ms=4, label=n5)
                plt.plot(x, c6, lw=1, c='darkgreen', marker='s', ms=4, label=n6)
                plt.plot(x, c7, lw=1, c='black', marker='s', ms=4, label=n7)
            plt.legend()
            ax.set_ylim(ax.get_ylim()[::-1])
            plt.title(str(draw_pic_num))
            plt.show()
        else:
            print("数据异常或不符合要求2")

if __name__ == '__main__':
    C0 = []
    C1 = []
    C2 = []
    C3 = []
    C4 = []
    C5 = []
    C6 = []
    C7 = []
    suject_num = -1
    suject = input("包括总分和所有科目，一共有多少个分数单元？请用阿拉伯数字输入")
    suject_num = int(suject)
    if(suject_num > 0):
        if(suject_num>=1):
            N1 = input("请问第一个科目是什么？(若为总分请输入「总计」)")
        if(suject_num>=2):
            N2 = input("请问第二个科目是什么？(若为总分请输入「总计」)")
        if (suject_num>=3):
            N3 = input("请问第三个科目是什么？(若为总分请输入「总计」)")
        if (suject_num>=4):
            N4 = input("请问第四个科目是什么？(若为总分请输入「总计」)")
        if (suject_num>=5):
            N5 = input("请问第五个科目是什么？(若为总分请输入「总计」)")
        if (suject_num>=6):
            N6 = input("请问第六个科目是什么？(若为总分请输入「总计」)")
        if (suject_num>=7):
            N7 = input("请问第七个科目是什么？(若为总分请输入「总计」)")

        path = input("请输入表格文件的路径：")
        sheet_Name = input("请输入子表格的名字： ")
        name_Line = int(input("从第几列开始是分数呢？请用阿拉伯数字输入： "))
        Width = int(input("一共有多少次考试的数据需要统计？请用阿拉伯数字输入： "))
        pic_time = int(input("请问您一共有多少位同学的数据想要打印？： "))

        if (suject_num == 1):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=None, n3=None, n4=None, n5=None, n6=None, n7=None,times=pic_time)
        elif (suject_num == 2):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=None, n4=None, n5=None, n6=None, n7=None,times=pic_time)
        elif (suject_num == 3):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=N3, n4=None, n5=None, n6=None, n7=None,times=pic_time)
        elif (suject_num == 4):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=N3, n4=N4, n5=None, n6=None, n7=None,times=pic_time)
        elif (suject_num == 5):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=N3, n4=N4, n5=N5, n6=None, n7=None,times=pic_time)
        elif (suject_num == 6):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=N3, n4=N4, n5=N5, n6=N6, n7=None,times=pic_time)
        elif (suject_num == 7):
            read_xlsx(c0=C0, c1=C1, c2=C2, c3=C3, c4=C4, c5=C5, c6=C6, c7=C7, xls_path=path
                      , sheet_name=sheet_Name, name_line=name_Line, nums_of_kinds=suject_num
                      , width=Width, n1=N1, n2=N2, n3=N3, n4=N4, n5=N5, n6=N6, n7=N7,times=pic_time)
    else:
        print("输入有误，请重新输入")
