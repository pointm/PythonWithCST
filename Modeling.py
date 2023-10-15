import win32com.client
import os


def quotes_to_list(text):
    '''
    这是一个将跨行字符串分行转化为数组的函数
    一般我在将VBA语句大卸八块的时候会用这个函数作为辅助
    '''
    lines = text.splitlines()
    quoted_lines = ["" + line.lstrip() + "" for line in lines]
    return quoted_lines


class COMWithHistory():
    mws = None

    Classname = 'COMWithHistory'

    def __init__(self, handle) -> None:
        self.mws = handle
        if self.mws == None:
            raise ('MWS未进行初始化，请重新进行mws初始化')
        else:
            print(self.Classname+'MWS'+'初始化成功')
            pass

    def AddToHistoryWithList(self, Tag, Command):
        '''
        添加到历史树
        Command只能是队列哦，不能传输裸字符串！
        适合在传递一小段一小段语句语句自定义程度较高的时候使用
        '''
        line_break = '\n'  # 这句话表示将每个单句用换行符链接起来
        Command = line_break.join(Command)  # 这句话表示将每个单句用换行符链接起来
        # AddToHistory ( string caption, string contents ) bool
        # 隶属于ProjectObject的方法，必须使用mws对象进行调用
        self.mws._FlagAsMethod("AddToHistory")
        self.mws.AddToHistory(Tag, Command)

    def AddToHistoryWithCommand(self, Tag, Command):
        '''
        添加到历史树
        可以传输字符串，适合整段整段懒得改的时候的语句使用
        '''
        # AddToHistory ( string caption, string contents ) bool
        # 隶属于ProjectObject的方法，必须使用mws对象进行调用
        self.mws._FlagAsMethod("AddToHistory")
        self.mws.AddToHistory(Tag, Command)


class Initial(COMWithHistory):
    cst = None
    mws = None

    def __init__(self, lable='New', ProjectName=' ') -> None:
        self.cst = self.OpenCST()
        if lable == 'New':
            handle = self.NewProject()
            self.BackGroundInitial(mws=handle)
            self.UnitInitial(mws=handle)
            self.BoundaryInitial(mws=handle)
        elif lable == 'Open':
            handle = self.OpenProject(ProjectName=ProjectName)

    def OpenCST(self):
        # 调用COM对象，COM对象的名字叫CSTStudio.Application
        self.cst = win32com.client.dynamic.Dispatch("CSTStudio.Application")
        return self.cst

    def NewProject(self):
        # 新建一个微波工作室，返回一个操控当前项目的COM组件
        self.cst.NewMWS()
        self.mws = self.cst.Active3D()
        return self.mws

    def OpenProject(self, ProjectName):
        # 打开一个项目，返回一个操控当前项目的COM组件
        self.cst.OpenFile(ProjectName)
        self.mws = self.cst.Active3D()
        return self.mws

    def BackGroundInitial(self, mws):
        sCommand = ['With Background',
                    '.ResetBackground',
                    '.Type "PEC"',
                    'End With']
        self.AddToHistoryWithList(
            mws=mws, Tag='Background Initial', Command=sCommand)

    def UnitInitial(self, mws):
        sCommand = ['With Units',
                    '.Geometry "mm"',
                    '.Frequency "ghz"',
                    '.Time "ns"',
                    'End With']
        self.AddToHistoryWithList(mws, Tag='Unit Initial', Command=sCommand)

    def BoundaryInitial(self, mws):
        sCommand = ['With Boundary',
                    '.Xmin "electric"',
                    '.Xmax "electric"',
                    '.Ymin "electric"',
                    '.Ymax "electric"',
                    '.Zmin "electric"',
                    '.Zmax "electric"',
                    '.Xsymmetry "none"',
                    '.Ysymmetry "none"',
                    '.Zsymmetry "none"',
                    'End With']
        self.AddToHistoryWithList(
            mws, Tag='Boundary Initial', Command=sCommand)

    def StoreParameter(self, mws, parametername, parametervalue, description):
        sCommand = '''  
        MakeSureParameterExists("%s", "%f")
        SetParameterDescription  ( "%s", "%s" )
    ''' % (parametername, parametervalue, parametername, description)
        self.AddToHistoryWithCommand(
            mws=mws, Tag='Store Parameter %s' % parametername, Command=sCommand)


class Material(COMWithHistory):
    def __init__(self, handle) -> None:
        self.mws = handle
        sCommand = '''With Material 
     .Reset 
     .Name "Sapphire"
     .Folder ""
     .Rho "0.0"
     .ThermalType "Normal"
     .ThermalConductivity "0"
     .SpecificHeat "0", "J/K/kg"
     .DynamicViscosity "0"
     .Emissivity "0"
     .MetabolicRate "0.0"
     .VoxelConvection "0.0"
     .BloodFlow "0"
     .MechanicsType "Unused"
     .IntrinsicCarrierDensity "0"
     .FrqType "all"
     .Type "Normal"
     .MaterialUnit "Frequency", "GHz"
     .MaterialUnit "Geometry", "mm"
     .MaterialUnit "Time", "ns"
     .MaterialUnit "Temperature", "Kelvin"
     .Epsilon "9.4"
     .Mu "1"
     .Sigma "0"
     .TanD "0.0"
     .TanDFreq "0.0"
     .TanDGiven "False"
     .TanDModel "ConstTanD"
     .SetConstTanDStrategyEps "AutomaticOrder"
     .ConstTanDModelOrderEps "3"
     .DjordjevicSarkarUpperFreqEps "0"
     .SetElParametricConductivity "False"
     .ReferenceCoordSystem "Global"
     .CoordSystemType "Cartesian"
     .SigmaM "0"
     .TanDM "0.0"
     .TanDMFreq "0.0"
     .TanDMGiven "False"
     .TanDMModel "ConstTanD"
     .SetConstTanDStrategyMu "AutomaticOrder"
     .ConstTanDModelOrderMu "3"
     .DjordjevicSarkarUpperFreqMu "0"
     .SetMagParametricConductivity "False"
     .DispModelEps  "None"
     .DispModelMu "None"
     .DispersiveFittingSchemeEps "Nth Order"
     .MaximalOrderNthModelFitEps "10"
     .ErrorLimitNthModelFitEps "0.1"
     .UseOnlyDataInSimFreqRangeNthModelEps "False"
     .DispersiveFittingSchemeMu "Nth Order"
     .MaximalOrderNthModelFitMu "10"
     .ErrorLimitNthModelFitMu "0.1"
     .UseOnlyDataInSimFreqRangeNthModelMu "False"
     .UseGeneralDispersionEps "False"
     .UseGeneralDispersionMu "False"
     .NLAnisotropy "False"
     .NLAStackingFactor "1"
     .NLADirectionX "1"
     .NLADirectionY "0"
     .NLADirectionZ "0"
     .Colour "0", "1", "0" 
     .Wireframe "False" 
     .Reflection "False" 
     .Allowoutline "True" 
     .Transparentoutline "False" 
     .Transparency "0" 
     .Create
End With
'''
        self.AddToHistoryWithCommand(
            mws=handle, Tag='addSapphire', Command=sCommand)


class GeneralModel(COMWithHistory):
    Component = ''
    Name = ''
    Material = ''

    def __init__(self, handle) -> None:
        self.mws = handle

    def init(self):
        pass

    def create(self):
        pass


class Brick(GeneralModel):
    Xrange = [0, 1]
    Yrange = [0, 1]
    Zrange = [0, 1]
    Component = 'Hallo'
    Name = 'World'
    Material = 'PEC'
    Classname = 'Brick'

    def init(self, Component, Name, Material, Xrange, Yrange, Zrange):
        self.Component = Component
        self.Name = Name
        self.Material = Material
        self.Xrange = Xrange
        self.Yrange = Yrange
        self.Zrange = Zrange

    def create(self, Tag):
        Command = ['With Brick',
                   '.Reset',
                   '.Name "%s"' % self.Name,
                   '.Component "%s"' % self.Component,
                   '.Material "%s"' % self.Material,
                   '.Xrange "%s", "%s"' % (self.Xrange[0], self.Xrange[1]),
                   '.Yrange "%s", "%s"' % (self.Yrange[0], self.Yrange[1]),
                   '.Zrange "%s", "%s"' % (self.Zrange[0], self.Zrange[1]),
                   '.Create',
                   'End With']
        self.AddToHistoryWithList(self.mws, Tag, Command)


class Cylinder(GeneralModel):
    Material = "Vacuum"
    Innerradius = 0
    Outerradius = 0
    Xcenter = 0
    Ycenter = 0
    Zcenter = 0
    Xrange = [0, 0]
    Yrange = [0, 0]
    Zrange = [0, 0]
    Range = [0, 0]
    Segments = 0
    Axis = 'z'

    def init(self, Component, Name, Material, Axis, Innerradius, Outerradius, Xcenter, Ycenter, Zcenter, Range, Segments):
        self.Component = Component
        self.Name = Name
        self.Material = Material
        self.Innerradius = Innerradius
        self.Outerradius = Outerradius
        self.Xcenter = Xcenter
        self.Ycenter = Ycenter
        self.Zcenter = Zcenter
        self.Range = Range
        self.Segments = Segments
        self.Axis = Axis

    def create(self, Tag):
        sCommand = ['With Cylinder',
                    '.Reset',
                    '.Name ("%s")' % self.Name,
                    '.Component ("%s")' % self.Component,
                    '.Material ("%s")' % self.Material,
                    '.Axis ("%s")' % self.Axis,
                    '.Outerradius ("%s")' % self.Outerradius,
                    '.Innerradius ("%s")' % self.Innerradius,
                    '.Xcenter (%f)' % self.Xcenter,
                    '.Ycenter (%f)' % self.Xcenter,
                    '.Zcenter (%f)' % self.Xcenter,]
        if self.Axis == 'z':
            sCommand = sCommand+['.Zrange ("%s", "%s")' % (self.Range[0], self.Range[1]),
                                 '.Segments (%s)' % self.Segments,
                                 '.Create',
                                 'End With']
        elif self.Axis == 'y':
            sCommand = sCommand+['.Yrange ("%s", "%s")' % (self.Range[0], self.Range[1]),
                                 '.Segments (%s)' % self.Segments,
                                 '.Create',
                                 'End With']
        elif self.Axis == 'x':
            sCommand = sCommand+['.Xrange ("%s", "%s")' % (self.Range[0], self.Range[1]),
                                 '.Segments (%s)' % self.Segments,
                                 '.Create',
                                 'End With']
        self.AddToHistoryWithList(mws=mws, Tag=Tag, Command=sCommand)


def CstSaveAsProject(mws, projectName):
    mws._FlagAsMethod("SaveAs")
    mws.SaveAs(projectName, 'false')


if __name__ == "__main__":
    path = os.getcwd()  # 获取当前py文件所在文件夹路径，方便保存
    filename = 'Test.cst'  # 保存的文件的名称，要加后缀cst
    projectName = os.path.join(path, filename)

    # init = Initial(lable='Open', ProjectName=projectName)
    init = Initial()
    mws = init.mws
    CstSaveAsProject(mws, projectName)  # 在新建时候保存用
    SimulateFrequency = [8, 9]

    history = COMWithHistory(mws)
    history.AddToHistoryWithCommand(
        'SetFrequencyRange', 'Solver.FrequencyRange "%f", "%f"' % (SimulateFrequency[0], SimulateFrequency[1]))
