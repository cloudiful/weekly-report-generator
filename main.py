# Prerequisite
# 1. 安装 xlwings 和对于 excel 的插件
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

        # 模式：
        #   1-自动模式：即当前选中的文件是上周的周报
        #   2-手动模式：即当前选中文件需修改日期或名字
        self.mode = 1

        # 用户姓名
        self.name = ''

        # 重命名后保存的文件
        self.newFile = None

        # 读取的excel数据
        self.wb = None
        self.sheet = None

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

        # 获取文件名中的日期字符串并转成datetime
        startDate = datetime.datetime.strptime(str(self.file)[startIndex:startIndex + 6], '%y%m%d')
        endDate = datetime.datetime.strptime(str(self.file)[endIndex:endIndex + 6], '%y%m%d')
        newStartDate = startDate + datetime.timedelta(days=7)
        newEndDate = endDate + datetime.timedelta(days=7)

        # 获取名字位置，即第二次出现'-'后面和文件后缀名前面
        originalName = str(self.file)[nameIndex:len(str(self.file)) - len(str(self.file.suffix))]
        newName = originalName

        if self.mode == '2':
            newName = input('改变名字（默认不修改）：')
            if newName == '':
                newName = originalName

        self.newFile = '周报（' + newStartDate.strftime('%y%m%d') + '-' + newEndDate.strftime(
            '%y%m%d') + '）- ' + newName + '.xlsx'
        print('新文件名：', self.newFile)

        shutil.copy2(self.file, self.newFile)
        print('新文件保存成功！')

    def changeDate(self):
        """变更excel中的文件"""
        print('- - - - - - - - - -\n'
              '正在进行日期变更')
        self.wb = xw.Book(self.newFile)
        self.sheet = self.wb.sheets[0]
        # 时间格式
        for index, date in enumerate(self.sheet['A1:H20']):
            if type(date.value) == datetime.datetime:
                date.value += datetime.timedelta(days=7)
        # 标题
        currentDay = datetime.datetime.today()
        currentWeekday = datetime.datetime.today().weekday()
        currentYear = datetime.datetime.today().strftime('%Y')
        print(currentDay - datetime.timedelta(days=currentWeekday))
        newStartDate = currentDay + datetime.timedelta(days=-currentWeekday)
        newEndDate = currentDay + datetime.timedelta(days=-currentWeekday+4)

        newTitle = currentYear + '年工作周报（' + datetime.datetime.strftime(newStartDate, '%m.%d') + '-' + datetime.datetime.strftime(newEndDate, '%m.%d') + '）'

        self.sheet['B1'].value = newTitle

        print('已完成日期变更！')

    def moveTexts(self):
        print('正在进行文本搬移')

        # 将本周工作移动到上周
        self.sheet['C16'].value = self.sheet['C4'].value
        self.sheet['C17'].value = self.sheet['C5'].value
        self.sheet['C18'].value = self.sheet['C6'].value

        # 将下周工作移动到本周
        self.sheet['C4'].value = self.sheet['C10'].value
        self.sheet['C5'].value = self.sheet['C11'].value
        self.sheet['C6'].value = self.sheet['C12'].value

        print('已完成文本搬移！')

        # 询问下周工作计划
        print('- - - - - - - - - -\n')
        self.sheet['C10'].value = input('请输入下周工作计划第1条:')
        self.sheet['C11'].value = input('请输入下周工作计划第2条:')
        self.sheet['C12'].value = input('请输入下周工作计划第3条:')


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
    fm.moveTexts()
