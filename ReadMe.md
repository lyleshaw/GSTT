# GSTT - Grades System Of TIANYI Test - [天一](http://tya.lhsvr.cn/index.php)成绩批量查询

## 缘起
你问我是不是这个系统的用户？我说我不是，将来报道出了偏差你们是要负责任的啊。  
只是有个[这个](http://tya.lhsvr.cn/index.php)系统的用户表示自己写了个脚本让我帮他改改。于是我就改了，顺便依照那人的要求加了一堆功能进去。  
由那位小伙伴起名，此repository名为**GSTT**。  

## 配置
请用文本编辑器打开`setting.json`。你将会看到如下内容：
```
{
    "wen_ke":[11,12,13,14,16],
    "li_ke":[1,2,3,4,5,6,7,8,9,10,15,17,18],
    "division": 1,
    "rank": 1,
    "kaohao":[1012180101,1012181850],
    "xm_name":"天一xxxxxxx"
}
```
现在说明如何配置。请勿复制`#`和之后的注解到`setting.json`中。
```
{
    "wen_ke":[11,12,13,14,16], # 方括号内填写文科班级，一个数字一个英文逗号，最后一个数字后没有英文逗号
    "li_ke":[1,2,3,4,5,6,7,8,9,10,15,17,18], # 方括号内填写理科班级，一个数字一个英文逗号，最后一个数字后没有英文逗号
    "division": 1, # 是否区分文理。是，填写1；否，填写0
    "rank": 1, # 成绩抓取后是否排名。是，填写1；否，填写0
    "kaohao":[1012180101,1012181850], # 方括号内填写初始考号，最后考号
    "xm_name":"天一xxxxxxx" # 引号内填写http://tya.lhsvr.cn/index.php 上'请选择考试项目'中你要抓取选项的文本
}
```

## 使用
### Linux
作为Linux用户你们一定明白我写了什么。提示一下输出有中文，所以系统语言要支持中文才是。

### Windows
（对于懂Python的用户）
想必你们已经安装了Python。在下使用Python 3.5.3。理论上兼容Python 3.X。
提示：记得pip安装pandas、openpyxl 、bs4。

（对于普通吃瓜群众）
[下载](https://github.com/Foldblade/GSTT/releases/download/1.0/GSTT.zip) ，参照之前的配置修改`setting.json`（记得保存）  
双击`GSTT.exe`运行。  
喝茶。看到卡住了敲下回车。  
继续喝茶。  
黑框框消失了、运行目录出现`output.xlsx`即已经成功。
