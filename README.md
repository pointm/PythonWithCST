## PythonWithCST
This is a repository for modeling using Python to manipulate the CST. 

这是一个用`Python`操纵`CST`进行建模的库。（没错都是机翻）

在进行测试的时候，`Python`的版本是3.11.4，`CST`的版本是2022版

本仓库的灵感来源于[这个仓库](https://github.com/kaankvrck/Cst-Py-Api)，这个[B站教程](https://www.bilibili.com/video/BV1d7411K77o/?share_source=copy_web&vd_source=2c9beb25af00b14851dca086bf631efd)以及[这个CSDN教程](https://blog.csdn.net/weixin_52556029/article/details/126983128)

他们的代码都很有启发性，但是CSDN教程只能在`PYTHON`老版本（具体来说，是`python`的3.6到3.9版本）才能运行，B站的教程能在高版本`python`运行，可以实现参数化建模，但是在进行扫参等操作的时候还是需要借助上面所提到的GitHub库里面提到的COM控件的方法进行操作。

This repository was inspired by [ this repository ](https://github.com/kaankvrck/Cst-Py-Api) , this [bilibili tutorial](https://www.bilibili.com/video/BV1d7411K77o/?share_source=copy_web&vd_source=2c9beb25af00b14851dca086bf631efd) , and [ this CSDN tutorial ](https://blog.CSDN.net/weixin_52556029/article/details/126983128) 

Their code is very enlightening, but CSDN tutorials can only run in older versions of Python (specifically, Python 3.6 to 3.9) , and bilibili tutorials can run in higher versions of Python, parameterized modeling can be achieved, but it is necessary to use the above-mentioned GitHub library in order to scan parameters and other operations of the COM control method.

## 使用(Usage)
请在运行之前确保自己的CST与Python已经安装正确~

Please make sure you have CST and Python installed correctly before you run it~!

现在集成的代码与`VBA`语句比较少，直接运行`ExampleOfMicrowaveWindow.py`模块即可预览里面的一个样例，具体的作用是会在文件当前的路径生成一个微波窗并且将里面的S参数绘制出图形。如果想要使用里面的类或者方法的话，直接将文件与`Modeling.py`放在同一个子文件夹里，然后`from Modeling import *`就可以了。

Now that the integrated code has fewer `VBA` statements, you can run the `ExampleOfMicrowaveWindow.py` module directly to preview one of its samples, which will generate a microwave window at the current path of the file and draw a graph with the S parameter inside. If you want to use any of the classes or methods in the module, place the file in the same subfolder as `Modeling.py`, and then `from Modeling import *`.

如何创建一个cst的文件和如何导入参数的话，直接参考`ExampleOfWaveGuide.py`文件就行啦。当然后续的话我也会考虑把这些比较重要的功能的使用方式添加到`README.md`里面的。

For how to create a cst file and how to import parameters, just refer to the `main` file in `Modeling.py`. Of course, I will consider adding the usage of these important functions to `README.md`.

如果有什么想要加的内容的话，欢迎提意见或者创建新的仓库分支！

If there's anything you'd like to add, feel free to comment or create a new branch of the repository!

## 现阶段BUG(Current Bugs)
- 现阶段使用里面的`Pick`类和`Port`类一同创建微波端口的话，打开创建好的文件可能会出现报错信息：`inconsistent model information please perform a update before start editing`。解决办法是前往`CST`的`History List`里面删除`Pick`与`Port`的代码，并且再去`CST`里面手动添加微波端口。`BUG`出现的具体的原因未知。

- At this stage, if you use the `Pick` and `Port` classes to create a microwave port, you may get an error message: `inconsistent model information please perform an update before starting editing` when you open the created file. The workaround is to go to the `History List` of the `CST` and delete the code for `Pick` and `Port`, and go to the `CST` to add the microwave port manually. The exact reason for the `bug` is unknown.

- 现阶段在脚本里面调用的求解器求解出来的结果与`CST`里面直接求解的结果有些许不同，换了两三种方法调用频域求解器都无法解决当前问题，后续可能会具体看看求解器方面的东西来解决这个`BUG`

- at this stage in the script to call the solver to solve the results and `CST ` inside the direct solution of the results are slightly different, change two or three ways to call the frequency domain solver can not solve the current problem, the follow-up may be specific look at the solver aspects of the things to solve this `bug `


## 后续想添加的一些功能(ToDoList)
尽量把建模部分做好吧。说实话，现在集成的`VBA`语句还是比较少，希望后续添加更多有用的东西。比如说通过扫描线按照轨迹进行扫描然后形成一个面之类的。

Let's try to get the modeling part right. To be honest, there are still relatively few `VBA` statements integrated, so I hope more useful things will be added. Like following a trajectory through a scan line and then forming a face or something like that.

## 注意(Attention)
- 1、直接使用以下的方法调用频域求解器的时候可能会出现报错。

  ```python
  init = Initial()
  mws = init.mws
  history = StructureMacros()
  sCommand='''
      With FDSolver
          FDSolver.Start
      End With
  '''
  history.AddToHistoryWithCommand('开始求解',sCommand)
  ```

  因为`FDSolver`的`Start`方法是一个`Control Macros`语句，而不是一个`Structure Macros`语句，严格上来说其并不应该被添加到`History List`之中。而应该使用`CST`的`COM`组件进行调用，具体的调用方式是：

  ```python
  init = Initial()
  mws = init.mws
  mws.FDSolver.Start#调用Start的时候请不要在后面加上括号哦，VBA里面原先的方法就没有括号
  ```

- 1. An error may be reported when calling the frequency domain solver directly using the following method.
  
  ```python
  init = Initial()
  mws = init.mws
  history = StructureMacros()
  sCommand='''
      With FDSolver
          FDSolver.Start
      End With
  '''
  history.AddToHistoryWithCommand('Start solving',sCommand)
  ```

  Because the `Start` method of `FDSolver` is a `Control Macros` statement, not a `Structure Macros` statement, it should not strictly speaking be added to the `History List`. Instead, it should be called using the `COM` component of `CST`, as follows:

  ```python
  init = Initial()
  mws = init.mws
  mws.FDSolver.Start#Please do not put parentheses after Start, the original method in VBA does not have parentheses.
  ```
