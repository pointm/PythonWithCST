import win32com.client
import os
import re


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
            Tag='Background Initial', Command=sCommand)

    def UnitInitial(self, mws):
        sCommand = ['With Units',
                    '.Geometry "mm"',
                    '.Frequency "ghz"',
                    '.Time "ns"',
                    'End With']
        self.AddToHistoryWithList(Tag='Unit Initial', Command=sCommand)

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
            Tag='Boundary Initial', Command=sCommand)

    def StoreParameters(self, parameters):
        sCommand = ''
        for parameter in parameters:
            sCommand = sCommand + '''  
            MakeSureParameterExists("{0}", {1})
            SetParameterDescription  ( "{2}", {3} )
        '''.format(parameter[0], parameter[1], parameter[0], parameter[2])
        # print(sCommand)
        self.AddToHistoryWithCommand(
            Tag='存储变量', Command=sCommand)

    def UseTemplate(self, Template, FrequencyRange):
        if Template == 'WaveGuide And Cavity Filter':
            sCommand = f'''
            'set the units
            With Units
                .Geometry "mm"
                .Frequency "GHz"
                .Voltage "V"
                .Resistance "Ohm"
                .Inductance "NanoH"
                .TemperatureUnit  "Kelvin"
                .Time "ns"
                .Current "A"
                .Conductance "Siemens"
                .Capacitance "PikoF"
            End With

            '----------------------------------------------------------------------------

            'set the frequency range
            Solver.FrequencyRange "{FrequencyRange[0]}", "{FrequencyRange[1]}"

            '----------------------------------------------------------------------------

            ' History:
            ' jei, vso 18-Jan-2012: ver1
            ' 28-Jan-2020: ver2

            ' boundaries
            With Boundary
                .Xmin "electric"
                .Xmax "electric"
                .Ymin "electric"
                .Ymax "electric"
                .Zmin "electric"
                .Zmax "electric"
            End With

            With Material
                .Reset
                .FrqType "all"
                .Type "Pec"
                .ChangeBackgroundMaterial
            End With

            With Mesh
                .MeshType "PBA"
                .SetCreator "High Frequency"
                .AutomeshRefineAtPecLines "True", "2"

                .UseRatioLimit "True"
                .RatioLimit "10"
                .LinesPerWavelength "20"
                .MinimumStepNumber "10"
                .Automesh "True"
            End With

            With MeshSettings
                .SetMeshType "Hex"
                .Set "StepsPerWaveNear", "13"
            End With

            ' solver - FD settings
            With FDSolver
                .Reset
                .Method "Tetrahedral Mesh" ' i.e. general purpose

                .AccuracyHex "1e-6"
                .AccuracyTet "1e-5"
                .AccuracySrf "1e-3"

                .SetUseFastResonantForSweepTet "False"

                .Type "Direct"
                .MeshAdaptionHex "False"
                .MeshAdaptionTet "True"

                .InterpolationSamples "5001"
            End With

            With MeshAdaption3D
                .SetType "HighFrequencyTet"
                .SetAdaptionStrategy "Energy"
                .MinPasses "3"
                .MaxPasses "10"
            End With

            With FDSolver
                .Method "Tetrahedral Mesh (MOR)"
                .HexMORSettings "", "5001"
            End With

            FDSolver.Method "Tetrahedral Mesh" ' i.e. general purpose

            '----------------------------------------------------------------------------

            With MeshSettings
                .SetMeshType "Tet"
                .Set "Version", 1%
            End With

            With Mesh
                .MeshType "Tetrahedral"
            End With

            'set the solver type
            ChangeSolverType("HF Frequency Domain")

            '----------------------------------------------------------------------------
    '''
        self.AddToHistoryWithCommand(
            'Template:WaveGuide And Cavity Filter', sCommand)


class Material(COMWithHistory):
    MaterialName = ''
    MaterialEpsilon = 0
    MaterialMu = 0

    def __init__(self, handle) -> None:
        self.mws = handle

    def materialinitial(self, Name, Epsilon, Mu):
        self.MaterialName = Name
        self.MaterialEpsilon = Epsilon
        self.MaterialMu = Mu
        return self

    def materialcreate(self):
        sCommand = f'''With Material 
     .Reset 
     .Name "{self.MaterialName}"
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
     .Epsilon "{self.MaterialEpsilon}"
     .Mu "{self.MaterialMu}"
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
            Tag='Add Material '+self.MaterialName, Command=sCommand)
        return self


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
        return self

    def create(self, Tag):
        Command = f'''With Brick
     .Reset 
     .Name "{self.Name}" 
     .Component "{self.Component}" 
     .Material "{self.Material}" 
     .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}" 
     .Yrange "{self.Yrange[0]}", "{self.Yrange[1]}"
     .Zrange "{self.Zrange[0]}", "{self.Zrange[1]}" 
     .Create
End With'''
        self.AddToHistoryWithCommand(Tag, Command)


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

    def init(self, Component, Name, Material, Axis, Innerradius, Outerradius, Center, Range, Segments=0):
        self.Component = Component
        self.Name = Name
        self.Material = Material
        self.Innerradius = Innerradius
        self.Outerradius = Outerradius
        self.Xcenter = Center[0]
        self.Ycenter = Center[1]
        self.Zcenter = Center[2]
        self.Range = Range
        self.Segments = Segments
        self.Axis = Axis
        return self

    def create(self, Tag):
        sCommand = f'''With Cylinder
    .Reset
    .Name ("{self.Name}")
    .Component ("{self.Component}")
    .Material ("{self.Material}")
    .Axis ("{self.Axis}")
    .Outerradius ("{self.Outerradius}")
    .Innerradius ("{self.Innerradius}")
    .Xcenter ("{self.Xcenter}")
    .Ycenter ("{self.Ycenter}")
    .Zcenter ("{self.Zcenter}")'''

        if self.Axis == 'z':
            sCommand = sCommand+f'''
    .Zrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
End With'''
        elif self.Axis == 'y':
            sCommand = sCommand+f'''
    .Yrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
End With'''
        elif self.Axis == 'x':
            sCommand = sCommand+f'''
    .Xrange ("{self.Range[0]}", "{self.Range[1]}")
    .Segments ("{self.Segments}")
    .Create
End With'''
        self.AddToHistoryWithCommand(Tag=Tag, Command=sCommand)


class Pick(COMWithHistory):
    def __init__(self, handle) -> None:
        self.mws = handle
        pass

    def PickCenterpointFromId(self, Tag, Component, Name, Id):
        sCommand = f'Pick.PickCenterpointFromId "{Component}:{Name}", "{Id}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def PickFaceFromId(self, Tag, Component, Name, Id):
        sCommand = f'Pick.PickFaceFromId "{Component}:{Name}", "{Id}"'
        self.AddToHistoryWithCommand(Tag, sCommand)


class WCS(COMWithHistory):
    def __init__(self, handle) -> None:
        self.mws = handle

    def AlignWCSWithSelectedPoint(self, Tag):
        self.AddToHistoryWithCommand(Tag, 'WCS.AlignWCSWithSelected "Point"')

    def ActivateWCSGlobal(self, Tag):
        self.AddToHistoryWithCommand(Tag, 'WCS.ActivateWCS "global"')


class Transform(COMWithHistory):
    def __init__(self, handle) -> None:
        self.mws = handle
        pass

    def MirrorTransForm(self, Tag, Component, Name, NormalVector, Copy):
        sCommand = f'''With Transform 
     .Reset 
     .Name "{Component}:{Name}" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "{NormalVector[0]}", "{NormalVector[1]}", "{NormalVector[2]}" 
     .MultipleObjects "{Copy}" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .AutoDestination "True" 
     .Transform "Shape", "Mirror" 
End With
'''
        self.AddToHistoryWithCommand(Tag, sCommand)


class Solid(COMWithHistory):
    def __init__(self, handle) -> None:
        self.mws = handle
    pass

    def Subtract(self, Tag, component1, name1, component2, name2):
        sCommand = f'Solid.Subtract "{component1}:{name1}", "{component2}:{name2}"'
        self.AddToHistoryWithCommand(Tag, sCommand)

    def Add(self, Tag, component1, name1, component2, name2):
        sCommand = f'Solid.Add "{component1}:{name1}", "{component2}:{name2}"'
        self.AddToHistoryWithCommand(Tag, sCommand)


class Port(COMWithHistory):
    PortNumber = 1
    NumberOfModes = 1
    Coordinates = 'Picks'
    Orientation = 'positive'
    PortOnBound = 'True'
    AdjustPolarization = 'False'
    Xrange = 0
    XrangeAdd = 0
    Yrange = 0
    YrangeAdd = 0
    Zrange = 0
    ZrangeAdd = 0

    def __init__(self, handle) -> None:
        self.mws = handle
    pass

    def init(self, Tag, Range, AddRange, **kwargs):
        self.Tag = Tag
        self.Xrange = Range[0]
        self.Yrange = Range[1]
        self.Zrange = Range[2]
        self.XrangeAdd = AddRange[0]
        self.YrangeAdd = AddRange[1]
        self.ZrangeAdd = AddRange[2]
        for key, value in kwargs.items():
            match key:
                case 'PortNumber':
                    self.PortNumber = value
                case 'NumberOfModes':
                    self.NumberOfModes = value
                case 'Coordinates':
                    self.Coordinates = value
                case 'Orientation':
                    self.Orientation = value
                case 'PortOnBound':
                    self.PortOnBound = value
                case 'AdjustPolarization':
                    self.AdjustPolarization = value

    def create(self):
        sCommand = f'''
    With Port 
        .Reset 
        .PortNumber "{self.PortNumber}" 
        .Label ""
        .Folder ""
        .NumberOfModes "{self.NumberOfModes}"
        .AdjustPolarization "{self.AdjustPolarization}"
        .PolarizationAngle "0.0"
        .ReferencePlaneDistance "0"
        .TextSize "50"
        .TextMaxLimit "0"
        .Coordinates "{self.Coordinates}"
        .Orientation "{self.Orientation}"
        .PortOnBound "{self.PortOnBound}"
        .ClipPickedPortToBound "False"
        .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}"
        .Xrange "{self.Xrange[0]}", "{self.Xrange[1]}"
        .Yrange "{self.Zrange[0]}", "{self.Zrange[1]}"
        .XrangeAdd "{self.XrangeAdd[0]}", "{self.XrangeAdd[1]}"
        .XrangeAdd "{self.XrangeAdd[0]}", "{self.XrangeAdd[1]}"
        .ZrangeAdd "{self.ZrangeAdd[0]}", "{self.ZrangeAdd[1]}"
        .SingleEnded "False"
        .WaveguideMonitor "False"
        .Create 
    End With'''
        self.AddToHistoryWithCommand(
            'Add Port' + str(self.PortNumber), sCommand)


class Mesh(COMWithHistory):
    StepsPerWaveNear = 17
    StepsPerWaveFar = 10
    StepsPerBoxNear = 12
    StepsPerBoxFar = 10
    MeshType = "Tetrahedral"
    SetCreator = "High Frequency"

    def __init__(self, handle) -> None:
        self.mws = handle

    def init(self, StepsPerWaveNear, StepsPerWaveFar, StepsPerBoxNear, StepsPerBoxFar, **kwargs):
        self.StepsPerBoxNear = StepsPerBoxNear
        self.StepsPerBoxFar = StepsPerBoxFar
        self.StepsPerWaveNear = StepsPerWaveNear
        self.StepsPerWaveFar = StepsPerWaveFar
        for key, value in kwargs.items():
            match key:
                case 'MeshType':
                    self.MeshType = value
                case 'SetCreator':
                    self.SetCreator = value

    def MeshUpdate(self, Tag):
        sCommand = f'''
        With Mesh 
            .MeshType "{self.MeshType}" 
            .SetCreator "{self.SetCreator}"
        End With 
        With MeshSettings 
            'MAX CELL - WAVELENGTH REFINEMENT 
            .Set "StepsPerWaveNear", "{self.StepsPerWaveNear}" 
            .Set "StepsPerWaveFar", "{self.StepsPerWaveFar}" 
            .Set "PhaseErrorNear", "0.02" 
            .Set "PhaseErrorFar", "0.02" 
            .Set "CellsPerWavelengthPolicy", "cellsperwavelength" 
            'MAX CELL - GEOMETRY REFINEMENT 
            .Set "StepsPerBoxNear", "{self.StepsPerBoxNear}" 
            .Set "StepsPerBoxFar", "{self.StepsPerBoxFar}" 
            .Set "ModelBoxDescrNear", "maxedge" 
            .Set "ModelBoxDescrFar", "maxedge" 
            'MIN CELL 
            .Set "UseRatioLimit", "0" 
            .Set "RatioLimit", "100" 
            .Set "MinStep", "0" 
            'MESHING METHOD 
            .SetMeshType "Unstr" 
            .Set "Method", "0" 
        End With 
        With MeshSettings 
            .SetMeshType "Tet" 
            .Set "CurvatureOrder", "1" 
            .Set "CurvatureOrderPolicy", "automatic" 
            .Set "CurvRefinementControl", "NormalTolerance" 
            .Set "NormalTolerance", "22.5" 
            .Set "SrfMeshGradation", "1.5" 
            .Set "SrfMeshOptimization", "1" 
        End With 
        With MeshSettings 
            .SetMeshType "Unstr" 
            .Set "UseMaterials",  "1" 
            .Set "MoveMesh", "0" 
        End With 
        With MeshSettings 
            .SetMeshType "All" 
            .Set "AutomaticEdgeRefinement",  "0" 
        End With 
        With MeshSettings 
            .SetMeshType "Tet" 
            .Set "UseAnisoCurveRefinement", "1" 
            .Set "UseSameSrfAndVolMeshGradation", "1" 
            .Set "VolMeshGradation", "1.5" 
            .Set "VolMeshOptimization", "1" 
        End With 
        With MeshSettings 
            .SetMeshType "Unstr" 
            .Set "SmallFeatureSize", "0" 
            .Set "CoincidenceTolerance", "1e-06" 
            .Set "SelfIntersectionCheck", "1" 
            .Set "OptimizeForPlanarStructures", "0" 
        End With 
        With Mesh 
            .SetParallelMesherMode "Tet", "maximum" 
            .SetMaxParallelMesherThreads "Tet", "1" 
        End With
        '''
        sCommand = sCommand+'''
        With Mesh 
            .Update 
        End With
        '''
        sCommand = sCommand + '''   
        With FDSolver
            .Start
        End With'''
        self.AddToHistoryWithCommand(Tag, sCommand)


def CstSaveAsProject(mws, projectName):
    mws._FlagAsMethod("SaveAs")
    mws.SaveAs(projectName, 'false')


if __name__ == "__main__":
    path = os.path.dirname(os.path.abspath(__file__))  # 获取当前py文件所在文件夹路径，方便保存
    filename = 'Test.cst'  # 保存的文件的名称，要加后缀cst
    projectName = os.path.join(path, filename)

    init = Initial(lable='Open', ProjectName=projectName)
    # init = Initial()
    mws = init.mws
    # CstSaveAsProject(mws, projectName)  # 在新建时候保存用
    SimulateFrequency = [8, 9]
    # 使用模板来对项目进行初始化
    history = COMWithHistory(mws)
    init.UseTemplate(Template='WaveGuide And Cavity Filter',
                     FrequencyRange=SimulateFrequency)

    # 加载变量名
    parametersfilename = 'ParameterList.txt'
    parameterspath = os.path.join(path, parametersfilename)
    originalparameters = open(parameterspath).readlines()  # 按行读取文件

    parameters = []
    for item in originalparameters:
        item = item.replace('\n', '')  # 去除换行符
        item = re.sub("[= ]", '#', item)  # 将没用的符号置换成分隔符
        item = item.split('#')  # 按照分隔符分开整行的指令
        parameters.append(item)

    # 将处理好的变量存储到应用中
    init.StoreParameters(parameters)

    # 创建材料Sapphire蓝宝石
    sapphire = Material(mws)
    sapphire.materialinitial('Sapphire', 6.5, 1)
    sapphire.materialcreate()

    # 创建圆柱形窗片
    cylinderwindow = Cylinder(mws)
    cylinderwindow.init('Window', 'SapphireWindow', sapphire.MaterialName,
                        'z', 0, 'wr', [0, 0, 0], ['-wt/2', 'wt/2'])
    cylinderwindow.create('创建圆柱形蓝宝石窗片')

    # 选取圆柱形窗片中点，将坐标系进行位移
    pick = Pick(mws)
    pick.PickCenterpointFromId(
        '选取圆柱窗片中心点', cylinderwindow.Component, cylinderwindow.Name, 3)
    wcs = WCS(mws)
    wcs.AlignWCSWithSelectedPoint('将中心点移到圆柱窗片中心')

    # 进行脊波导的建模
    waveguide = Brick(mws)
    L = 10
    waveguide.init('WaveGuide', 'DRW', 'Vacuum', [
                   '-a/2', 'a/2'], ['-b/2', 'b/2'], [0, L])
    waveguide.create('创建双脊波导本体')

    cutoff = Brick(mws)
    cutoff.init('WaveGuide', 'cutoff', 'Vacuum', [
                '-d/2', 'd/2'], ['c/2', 'b/2'], [0, L])
    cutoff.create('创建被切除部分')

    trans = Transform(mws)
    trans.MirrorTransForm('镜像切除部分', cutoff.Component,
                          cutoff.Name, [0, -1, 0], True)

    # 切除波导冗余部位
    solid = Solid(mws)
    solid.Subtract('开始减去部位1', waveguide.Component, waveguide.Name,
                   cutoff.Component, cutoff.Name)
    solid.Subtract('开始减去部位2', waveguide.Component, waveguide.Name,
                   cutoff.Component, cutoff.Name+'_1')
    # trans.MirrorTransForm('镜像脊波导', waveguide.Component,
    #                       waveguide.Name, [0, 0, -1], True)

    # 补偿波导建模，顺便一提论文的这个部分有问题，具体的宽度我只能脑测了
    transportwaveguide = Brick(mws)
    transportwaveguide.init(waveguide.Component, 'TW', waveguide.Material, [
                            '-a/2', 'a/2'], ['-b/2*0.75', 'b/2*0.75'], [0, 't'])
    transportwaveguide.create('添加过渡波导')
    solid.Add('将脊波导与过渡波导相加', waveguide.Component,
              waveguide.Name, waveguide.Component, 'TW')

    # 创建全局坐标系，进行变换
    wcs.ActivateWCSGlobal('激活全局坐标系，准备变换')
    trans.MirrorTransForm('将创建完成的脊波导进行镜像', waveguide.Component,
                          waveguide.Name, [0, 0, -1], True)
    # 选取面，并且设置端口
    pick.PickFaceFromId('选取面1', waveguide.Component, waveguide.Name, 27)
    setport = Port(mws)
    setport.init('添加端口1', [['-a/2', 'a/2'], ['-b/2', 'b/2'],
                           ['wt/2+10', 'wt/2+10']], [[0, 0], [0, 0], [0, 0]], PortNumber=1)
    setport.create()

    pick.PickFaceFromId('选取面2', waveguide.Component, waveguide.Name + '_1', 27)
    setport.init('添加端口2', [['-a/2', 'a/2'], ['-b/2', 'b/2'],
                           ['-(wt/2+10)', '-(wt/2+10)']], [[0, 0], [0, 0], [0, 0]], PortNumber=2)
    setport.create()

    # 更新网格并且求解(我不会写网格更新（悲）)
    mesh = Mesh(mws)
    mesh.init(10, 5, 6, 5)
    mesh.MeshUpdate('网格更新')

    # 求解后处理

    pass
