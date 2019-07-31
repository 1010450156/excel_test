import xlwt
import time,os
import random
import logging
from datetime import date, timedelta

#日志打印
logger = logging.getLogger(__name__)
logger.setLevel(level=logging.DEBUG)
handler = logging.FileHandler("D:\\测试任务\\week\\code\\log\\test_log.log")
handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


class file_process():
    # 获取某一页sheet对象
    def create_excel_file(self):
        da = time.strftime("%Y-%m-%d", time.localtime())
        self.da_a = str(da.split('/')[0])
        filename = (self.da_a)+".xlsx"
        excel_path = os.path.dirname(os.path.abspath('.')) + '\data'
        self.ecl =os.path.join(excel_path,filename)

        # 如果存在filename 则先删掉且记录日志
        filepath = excel_path + '/' + filename
        if os.path.exists(filepath):
            os.remove(filepath)
            logger.info("已存在当前日期文件，删除此文件:{}".format(filepath))
        # 删除历史文件
        self.clear_file()


        style = xlwt.XFStyle()  # 格式信息
        font = xlwt.Font()  # 字体基本设置
        font.name = 'Times New Roman'
        font.bold = True  # 黑体
        font.color = 'black'
        style.font = font

        book = xlwt.Workbook(encoding = 'utf-8')
        self.sheet1 = book.add_sheet('自动化1',cell_overwrite_ok = True)
        self.title = [u'账号',u'密码']
        self.write_data_one(self.title,0,line=1)#标题内容和格式写入
        # self.a = xlwt.Formula('AVERAGE(A2,A1000)')
        for i in range(2):
            self.sheet1.write(0, 0, self.title[i], style)
            self.sheet1.write(0, 1, self.title[i], style)
            self.sheet1.col(i).width = 5000
            self.sheet1.col(i).width = 5000

        self.sheet2 = book.add_sheet('自动化2', cell_overwrite_ok=True)
        self.titletwo = [u'平均值']
        self.sheet2.write(0, 1, xlwt.Formula('AVERAGE(自动化1!A2,自动化1!A1000)'))
        self.write_data_two(self.titletwo, 0, line=1)  # 标题内容和格式写入
        self.sheet2.write(0, 0, self.titletwo, style)
        self.sheet2.col(1).width = 5000


        # 数据导入的条数
        data_num = 1000
        # 账号写入
        i = self.title_local(u'账号')
        brrower_name = self.create_data_account(data_num)  # 获取5个随机字符
        self.write_data_one(brrower_name, i, line=0)
        #密码写入
        ff = self.title_local(u'密码')
        loan_institutions = self.create_data_password(data_num)
        self.write_data_one(loan_institutions, ff, line=0)

        book.save(self.ecl)
        return filename

    def write_data_one(self,lst,num,line=0):
        if line == 0:
            for i, item in enumerate(lst):
                self.sheet1.write(i + 1, num, item)
        else:
            for i, item in enumerate(lst):
                self.sheet1.write(num, i, item)

    def write_data_two(self, lst, num, line=0):
        if line == 0:
            for i, item in enumerate(lst):
                self.sheet2.write(i + 1, num, item)
        else:
            for i, item in enumerate(lst):
                self.sheet2.write(num, i, item)


    def title_local(self,str):
        for i, item in enumerate(self.title):
            if item == str:
                return i

    def create_data_account(self,number):
        lrst = []
        la = 8451252630000
        lb = 1
        for a in range(1, number+1):
            ha = la + lb
            lb = lb+1
            lrst.append(ha)
        return lrst

    def create_data_password(self,data_num):
        loan_institutions = []
        self.loan_institution = ['Faxuan.%1234']
        for b in range(1,data_num+1):
            ff = random.choice(self.loan_institution)
            loan_institutions.append(ff)
        return loan_institutions

    #将不是今天生成的文件删除
    def clear_file(self):
        # 删除一天前的文件
        da_a = str(date.today() + timedelta(days = -2))
        path = os.path.dirname(os.path.abspath('.')) + '\data'
        lrst = []
        for file in os.listdir(path):
            if da_a not in file:
                lrst.append(file)
        if len(lrst) == 0:
            logger.info("没有需要删除的历史文件")
        else:
            for file in lrst:
                os.remove(os.path.join(path,file))
                logger.info("删除历史文件{}".format(path + '/' + file))


if __name__ == '__main__':
    start = time.clock()
    a=file_process()
    a.create_excel_file()
    end = time.clock()
    T = end - start
    T = "%.2f" % T
    print("程序运行时间是：" + T + "秒" )
    logger.info("程序运行时间是：" + T + "秒" )
