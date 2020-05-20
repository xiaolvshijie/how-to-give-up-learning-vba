
#  📺Part.1.B.start-coding（开始编程）

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180557.png)

## 一、在哪里开始写代码？

Excel中编辑代码的区域默认是关闭的，基础用户默认是无法编辑VBA代码的，需要勾选「开发工具」选项卡，才可以打开VB代码编辑器，或者直接按下快捷键「ALT+F11」

> 🎬操作步骤
> 1. 文件 > 选项 > 自定义功能区 >「开发工具」选项卡 > Visual Basic
> 2. 快捷键：ALT+F11
> 3. 其他

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180559.gif)


## 二、设置宏安全性

默认情况下，为防止来源不明的工作簿自带宏自动运行，Excel会禁用宏的运行。为了运行自己录制的宏，按如下步骤设置宏安全性。先点击开发工具选项卡里，「宏安全性」命令。在弹出的设置菜单中，按如下方式设置。

在学习集中VBA的过程中也可以选择「启用所有宏」。

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180603.png)

## 三、设置Visual Basic编辑器

将Visual Basic编辑器按照以下设置，在**视图菜单**打开代码窗口、立即窗口、本地窗口、工程资源管理器和属性窗口，便于之后编辑代码


![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180606.png)

选择适合的代码字体和字号，推荐 **微软雅黑+12号**


![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180611.png)

## 四、尝试写第一段代码

在Visual Basic编辑器中插入模块，将以下代码输入至代码窗口，体验敲代码的感觉，**不推荐直接复制**，单击「▶」或者按快捷键`F5`运行代码。

~~~Visual Basic
' 代码注释：在弹出的窗口显示“我在学习VBA”
Sub start()
    MsgBox "我在学习VBA"
End Sub
~~~

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180613.gif)

## 五、怎么保存包含代码的Excel工作簿

写好代码之后，需要将Excel工作簿另存为（F12）**.xlsm、.xlsb、.xls**三种格式，这样才能保存写好的代码，如果用 .xlsx 格式保存Excel工作簿，那么将会丢失已经写好的代码‼。

> 推荐首先新建一个 .xlsm 格式的Excel工作簿，然后再开始写代码，避免保存错误。


![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520180628.png)


##  😈本章作业


参照本章示例，设置你的Visual Basic编辑器，尝试运行第一段代码，最后把包含这段代码的Excel工作簿用合适的格式保存。


![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/公众号名片.jpg)

- - -

⬆上一节：[Part.1.A.what-is-vba（**认识VBA**）](Part.1.A.what-is-vba.md)

⬇下一节：[Part.1.C.macro-v2.0（**录制宏**）](Part.1.C.macro-v2.0.md)