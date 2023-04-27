# weekly-report-generator
a generator for bullshit weekly report
本软件要求你的电脑上已经装有Microsoft Excel软件

Prerequisite:
1. 安装 xlwings 和 excel 的插件
   （参考 https://docs.xlwings.org/zh_TW/latest/installation.html#installation）
      省流版本：先 pip install xlwings 或 conda install xlwings
              再 xlwings addin install
              
Usage:
1.本脚本目前尚未完全实现格式自适应，请查看自己的周报格式和周报模版是否一致
2.将需要修改的周报文件复制到脚本同目录下
3.运行py文件，脚本将新建excel文件，并自动打开excel软件
4.本脚本根据目前时间自动修改日期，并将本周工作内容移到上周，将下周工作内容移到本周，并根据需要修改执行人姓名
