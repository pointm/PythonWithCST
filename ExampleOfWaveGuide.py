from Modeling import *

path = os.path.dirname(os.path.abspath(__file__))  # 获取当前py文件所在文件夹路径，方便保存
filename = 'Test.cst'  # 保存的文件的名称，要加后缀cst
projectName = os.path.join(path, filename)

# init = Initial(lable='Open', ProjectName=projectName)
init = Initial()
mws = init.mws
cst = init.cst
CstSaveAsProject(mws, projectName)  # 在新建时候保存用
SimulateFrequency = [8, 9]
# 使用模板来对项目进行初始化
history = StructureMacros(mws)
init.UseTemplate(Template='WaveGuide And Cavity Filter',
                 FrequencyRange=SimulateFrequency)

# 加载变量名
parametersfilename = 'ParameterList.txt'
parameterspath = os.path.join(path, parametersfilename)
# 按行读取文件，有可能会出现编码错误，出现的话记得在open的后面加上编码(encode='utf-8')当然没出现的话就不管了
originalparameters = open(parameterspath).readlines()

parameters = []
for item in originalparameters:
    item = item.replace('\n', '')  # 去除换行符
    item = re.sub("[= ]", '#', item)  # 将没用的符号批量置换成分隔符
    item = item.split('#')  # 按照分隔符分开整行
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
setport.init([['-a/2', 'a/2'], ['-b/2', 'b/2'],
              ['wt/2+10', 'wt/2+10']], PortNumber=1)
setport.create('添加端口1')

pick.PickFaceFromId('选取面2', waveguide.Component, waveguide.Name + '_1', 27)
setport.init([['-a/2', 'a/2'], ['-b/2', 'b/2'],
              ['-(wt/2+10)', '-(wt/2+10)']], PortNumber=2)
setport.create('添加端口2')

# 更新网格并且求解(我不会写网格更新（悲）)
mesh = Mesh(mws)
mesh.init(10, 5, 6, 5)
mesh.MeshUpdate('网格更新')

# 进行求解(我也不会写求解器(悲))
solver = Solver(mws)
solver.FDSolver()
# 求解S参数后处理
postprocess = PostProcessingItems(mws)
resultdatas, Frequencyseries, SREALseries, SIMAGEseries = postprocess.GetSparametersinRunID(
    ResultTag='S12')
Sdbs = []
for runid, Sreal in enumerate(SREALseries):
    Sdb = []
    for index, Sparameter in enumerate(Sreal):
        Sparameterdb = 20 * \
            math.log10(
                abs(complex(Sreal[index], SIMAGEseries[runid][index])))
        Sdb.append(Sparameterdb)
    # plt.figure(runid)  # 注释了的话那就在同一张图上
    plt.plot(Frequencyseries[runid], Sdb)
    plt.xlabel(resultdatas[runid].GetXlabel)
    plt.ylabel('Magnitude in dB')
    plt.title('S11, Current RunId:' + str(runid))
    # plt.show()
    Sdbs.append(Sdb)
plt.show()

pathofselected = postprocess.GetSelectedTreeItem()
pass
