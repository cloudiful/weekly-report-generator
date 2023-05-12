# Prerequisite:
# 1. 安装 xlwings 和 excel 的插件
#    （参考 https://docs.xlwings.org/zh_TW/latest/installation.html#installation）
#    先 pip install xlwings 或 conda install xlwings
#    再 xlwings addin install

import datetime
import shutil
import xlwings as xw

from pathlib import Path


class FileModify:

    def __init__(self):
        # 文件路径设置为当前目录
        self.path = Path('.')

        # 当前目录下所有文件
        self.files = None

        # 用户选择的文件
        self.file = None

        # 当前日期信息
        self.currentDate = None
        self.currentWeekday = None
        self.currentYear = None
        # 预期输出的日期
        self.startDate = None
        self.endDate = None

        # 用户姓名
        self.name = ''

        # 重命名后保存的文件
        self.newFile = None

        # 读取的excel数据
        self.wb = None
        self.sheet = None

        self.getTime()

    def showFiles(self):
        """显示当前目录下的所有excel文件"""
        self.files = [e for e in self.path.iterdir() if e.is_file() and (e.suffix == '.xlsx' or e.suffix == '.xls')]
        print('- - - - - - - - - -')
        for index, file in enumerate(self.files):
            print(index, ' - ', file)
        return self.files

    def chooseFile(self):
        """如果当前目录下有多个excel文件，则询问用户选择哪个文件"""
        i = input('- - - - - - - - - -\n'
                  "请选择上周的周报文件：")

        while i == '' or int(i) < 0 or int(i) > len(self.files) - 1:
            print('\n请输入正确的文件序号！\n')
            self.showFiles()
            i = input("请选择上周的周报文件：")

        self.file = self.files[int(i)]

    def getTime(self):
        self.currentDate = datetime.datetime.today()
        self.currentWeekday = datetime.datetime.today().weekday()
        self.currentYear = datetime.datetime.today().strftime('%Y')
        self.startDate = self.currentDate + datetime.timedelta(days=-self.currentWeekday)
        self.endDate = self.currentDate + datetime.timedelta(days=-self.currentWeekday + 4)

    def modifyFileName(self):
        """更新文件名中的日期和姓名"""
        print('- - - - - - - - - -\n'
              '原文件名：', str(self.file))
        # 日期在文件名中出现的位置 例如 "周报（230417-230421）"中日期在文件名中的出现的位置为 3
        # 左括号右一位为开始日期
        startIndex = str(self.file).find('（') + 1
        # - 右一位为结束日期
        endIndex = str(self.file).find('-') + 1
        # 右括号右三位为名字
        nameIndex = str(self.file).find('）') + 3

        # 获取名字位置，即第二次出现'-'后面和文件后缀名前面
        originalName = str(self.file)[nameIndex:len(str(self.file)) - len(str(self.file.suffix))]

        self.name = input('输入执行人姓名（直接回车即默认不修改）：')
        if self.name == '':
            self.name = originalName

        self.newFile = '周报（' + self.startDate.strftime('%y%m%d') + '-' + self.endDate.strftime(
            '%y%m%d') + '）- ' + self.name + '.xlsx'

        if str(self.file) == str(self.newFile):
            print('- - - - - - - - - -\n'
                  '当前文件名不需要修改，将直接对当前文件进行修改')
        else:
            print('- - - - - - - - - -\n'
                  '新文件名：', self.newFile)
            shutil.copy2(self.file, self.newFile)
            print('新文件保存成功！')

    def changeDate(self):
        """变更excel中的文件"""
        print('- - - - - - - - - -\n'
              '正在进行日期变更')
        self.wb = xw.Book(self.newFile)
        self.sheet = self.wb.sheets[0]
        # 本周时间
        for index, date in enumerate(self.sheet['D5:D7']):
            date.value = self.startDate
        for index, date in enumerate(self.sheet['E5:E7']):
            date.value = self.endDate
        # 下周时间
        for index, date in enumerate(self.sheet['D11:D13']):
            date.value = self.startDate + datetime.timedelta(days=7)
        for index, date in enumerate(self.sheet['E11:E13']):
            date.value = self.endDate + datetime.timedelta(days=7)

        # 上周时间
        for index, date in enumerate(self.sheet['D17:D19']):
            date.value = self.startDate - datetime.timedelta(days=7)
        for index, date in enumerate(self.sheet['E17:E19']):
            date.value = self.endDate - datetime.timedelta(days=7)

        # 标题

        newTitle = self.currentYear + '年工作周报（' + datetime.datetime.strftime(self.startDate,
                                                                                 '%m.%d') + '-' + datetime.datetime.strftime(
            self.endDate, '%m.%d') + '）'

        self.sheet['B1'].value = newTitle

        print('已完成日期变更！')

    def changeTexts(self):
        print('- - - - - - - - - -\n'
              '正在进行工作内容修改')

        # 添加序号
        self.sheet['B5,B11,B17'].value = '1'
        self.sheet['B6,B12,B18'].value = '2'
        self.sheet['B7,B13,B19'].value = '3'

        # 添加本周工作内容完成情况
        self.sheet['F5:F7'].value = '完成'

        # 将本周工作移动到上周或要求输入上周工作内容
        for i, x in enumerate(self.sheet['C5:C7'].value):
            if x is None:
                self.sheet['C' + str(17 + i)].value = input('请输入上周工作内容第' + str(i + 1) + '条：')
            else:
                self.sheet['C' + str(17 + i)].value = self.sheet['C' + str(5 + i)].value

        # 将下周工作移动到本周或要求输入本周工作内容
        for i, x in enumerate(self.sheet['C11:C13'].value):
            if x is None:
                self.sheet['C' + str(5 + i)].value = input('请输入本周工作内容第' + str(i + 1) + '条：')
            else:
                self.sheet['C' + str(5 + i)].value = self.sheet['C' + str(11 + i)].value

        # 询问下周工作计划
        self.sheet['C11'].value = input('请输入下周工作计划第1条:')
        self.sheet['C12'].value = input('请输入下周工作计划第2条:')
        self.sheet['C13'].value = input('请输入下周工作计划第3条:')

        print('已完成工作内容修改！')

    def changeName(self):
        print('- - - - - - - - - -\n'
              '正在进行执行人修改')
        self.sheet['G5:G7'].value = self.name
        self.sheet['F11:F13'].value = self.name
        self.sheet['F17:F19'].value = self.name
        print('执行人修改完成！')

    def blankDetect(self):
        """最后阶段检测是否有空的工作内容，如果有则删除整行"""
        for i, x in enumerate(self.sheet['C5:C7'].value):
            if x is None:
                self.sheet['B' + str(5 + i) + ':G' + str(5 + i)].value = None
        for i, x in enumerate(self.sheet['C11:C13'].value):
            if x is None:
                self.sheet['B' + str(11 + i) + ':G' + str(11 + i)].value = None
        for i, x in enumerate(self.sheet['C17:C19'].value):
            if x is None:
                self.sheet['B' + str(17 + i) + ':G' + str(17 + i)].value = None


if __name__ == '__main__':
    # 初始化类
    fm = FileModify()

    # 显示同目录下的excel文件
    fm.showFiles()

    # 选择上周excel文件
    fm.chooseFile()

    # 修改文件名并另存为新excel
    fm.modifyFileName()

    # 时间加7天
    fm.changeDate()

    # 移动本周工作情况到上周 移动下周工作计划到本周工作情况
    fm.changeTexts()

    # 修改执行人姓名
    fm.changeName()

    # 空内容检测
    fm.blankDetect()

    print('所有操作均已完成，请在Excel软件中保存')
