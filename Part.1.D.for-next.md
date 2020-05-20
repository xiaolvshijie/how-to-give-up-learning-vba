#   📺Part.1.D.for-next（使用For循环）

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520181252.png)

## 一、上一节录制宏得到的代码

~~~vb
Sub 工资条()
'
' 工资条 宏
' 制作工资条
'
' 快捷键: Ctrl+q
'
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveCell.Rows("1:1").EntireRow.Select
End Sub
~~~

**Tips**：如果你的工作簿中没有找到上面这段代码，可能是你保存Excel时没有选择正确的格式，使用`.xlsm`格式保存含有宏的工作簿。

这段代码就是上一节中录制宏得到的代码，这一段代码可以实现自动插入一行工资条标题。如果有7条工资信息，除了第一行，那么就需要插人六行工资条标题，按照上节课的做法我们只需要连续运行6次宏即可，但是实际工作中工资条可能远不止6条，那么我们如何解决这个问题？

## 二、加上For循环后的工资条

为解决上述多次重复运行同一段宏代码的问题，我们可以使用`For···Next`循环语法，将上节中录制宏得到的代码用`For···Next`循环语法嵌套，现在我们不需要掌握`For···Next`循环语法的详细概念，**先实践，再概念**。由于上一节中我们需要连续运行6次宏，即需要对这段代码循环6次，所以嵌套上 `for i = 1 to 6···next`，如果需要循环100次，那么嵌套上 `for i = 1 to 100···next`。加上For循环后，**选中第一行**，开始执行宏，一个2.0版本的工资条工具就制作完成了，之后还有3.0版本，如下所示。

**Tips**：选中代码行，按**`Tab`**键可以批量缩进代码，按**`Shift+Tab`**键可以批量取消缩进代码。

~~~vb
Sub 工资条2()
'
' 工资条 宏
' 制作工资条
'
' 快捷键: Ctrl+q
'	
    Dim i 
    for i = 1 to 6
        ActiveCell.Rows("1:1").EntireRow.Select
        Selection.Copy
        ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
        Selection.Insert Shift:=xlDown
        ActiveCell.Rows("1:1").EntireRow.Select
    next
End Sub
~~~

![](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520181257.gif)

## 三、本节小结😵

### （一）子过程

```vb
Sub 工资条() '过程开始，过程的名称为“工资条”
	······  '代码正文
End Sub  '结束过程
```

VBA过程包括三类：子过程、函数过程和属性过程。上面就是子过程的**标准格式**，标志是以Sub开头，所有录制宏产生的过程都是子过程。使用VBA时，基本都是使用子过程，其他类型过程用的较少，暂时不提，减少点理解负担。

**Tips**：代码后用英文单引号`'`插入代码注释，不会影响程序运行，只是为了便于理解。

### （二）变量和For循环

```vb
Dim i  '定义变量
```

变量没有固定的值，可以随时根据需求赋予新值。类似于数学中的**设未知数X解方程**，这个X就可以理解为变量。

```vb
For i= 1 to 10  'For循环开始
	······
Next  '继续循环至结束
```

`For···Next`语句表示以指定次数来重复执行一组语句,一般需要配合变量使用，变量对于编程来说都是极为重要的概念。

```vb
Sub 使用变量()
    Dim i
    For i = 1 To 6
        Debug.Print i
    Next
End Sub
```

运行上面这段代码后，我们观察**立即窗口**可以看到连续从1输出到6，由此可以看出For循环语法的作用，从`i=1`开始循环执行`Debug.Print i`这一代码到`i=6`为止。

![1D-2](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/20200520181302.png)

基于此，我们才可以顺利做到隔行插入工资条标题，已经其他更多操作······


##  😈本章作业


下载本章的[**示例文件**（Sample_file/Part.1.D.for-next）](Sample_file/Part.1.D.for-next)

参照本章示例，将工资条还原成原始表格，也就是说，用For循环和录制宏将多余的工资条标题删除，只保留第一行。

![1D-2](https://xlsj.oss-cn-beijing.aliyuncs.com/wwc_pic/公众号名片.jpg)

- - -

⬆上一节：[Part.1.C.macro-v2.0（**录制宏**）](Part.1.C.macro-v2.0.md)

⬇下一节：[Part.1.C.macro-v2.0（**录制宏**）](Part.1.C.macro-v2.0.md)