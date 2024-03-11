from Modeling import *

path = os.path.dirname(
    os.path.abspath(__file__)
)  # 获取当前py文件所在文件夹路径，方便保存
filename = "MMW.cst"  # 保存的文件的名称，要加后缀cst
projectName = os.path.join(path, filename)
init = Initial(lable="New", ProjectName=projectName)

mws = init.mws
cst = init.cst

PECCUBE = Brick(
    handle=mws,
    Tag="PECCUBE",
    Component="TEST",
    Name="TestCube1",
    Material="PEC",
    Xrange=[0, 1],
    Yrange=[0, 1],
    Zrange=[0, 1],
)

pass
# cst.Quit()

VACUUMCUBE = Brick(
    handle=mws,
    Tag="VACUUMCUBE",
    Component="TEST",
    Name="TestCube2",
    Material="Vacuum",
    Xrange=[-1, 0],
    Yrange=[-1, 0],
    Zrange=[-1, 0],
)
pass


cst.Quit()
