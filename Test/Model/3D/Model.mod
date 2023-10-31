'# MWS Version: Version 2022.5 - Jun 03 2022 - ACIS 31.0.1 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 8 fmax = 9
'# created = '[VERSION]2022.5|31.0.1|20220603[/VERSION]


'@ Background Initial

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Background
.ResetBackground
.Type "PEC"
End With

'@ Unit Initial

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Units
.Geometry "mm"
.Frequency "ghz"
.Time "ns"
End With

'@ Boundary Initial

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Boundary
.Xmin "electric"
.Xmax "electric"
.Ymin "electric"
.Ymax "electric"
.Zmin "electric"
.Zmax "electric"
.Xsymmetry "none"
.Ysymmetry "none"
.Zsymmetry "none"
End With

'@ Template:WaveGuide And Cavity Filter

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
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
            Solver.FrequencyRange "8", "9"

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

'@ 存储变量

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
MakeSureParameterExists("a", "18.29")
            SetParameterDescription  ( "a", "a" )
          
            MakeSureParameterExists("b", "8.15")
            SetParameterDescription  ( "b", "b" )
          
            MakeSureParameterExists("c", "4.39")
            SetParameterDescription  ( "c", "c" )
          
            MakeSureParameterExists("d", "2.57")
            SetParameterDescription  ( "d", "d" )
          
            MakeSureParameterExists("s", "4.35")
            SetParameterDescription  ( "s", "s" )
          
            MakeSureParameterExists("wt", "0.15785191304674")
            SetParameterDescription  ( "wt", "wt" )
          
            MakeSureParameterExists("wr", "6.3620038786882")
            SetParameterDescription  ( "wr", "wr" )
          
            MakeSureParameterExists("trh", "0.19296535797318")
            SetParameterDescription  ( "trh", "trh" )
          
            MakeSureParameterExists("ta", "a")
            SetParameterDescription  ( "ta", "ta" )
          
            MakeSureParameterExists("tb", "2.7735691617668")
            SetParameterDescription  ( "tb", "tb" )

'@ Add Material Sapphire

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Material 
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
     .Epsilon "6.5"
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

'@ 创建圆柱形蓝宝石窗片

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Cylinder
    .Reset
    .Name ("SapphireWindow")
    .Component ("Window")
    .Material ("Sapphire")
    .Axis ("z")
    .Outerradius ("wr")
    .Innerradius ("0")
    .Xcenter ("0")
    .Ycenter ("0")
    .Zcenter ("0")
    .Zrange ("-wt/2", "wt/2")
    .Segments ("0")
    .Create
    End With

'@ 选取圆柱窗片中心点

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Pick.PickCenterpointFromId "Window:SapphireWindow", "3"

'@ 将中心点移到圆柱窗片中心

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
WCS.AlignWCSWithSelected "Point"

'@ 创建双脊波导 DoubleRidgeWaveGuide 本体

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Brick
     .Reset 
     .Name "DRW1" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-a/2", "a/2" 
     .Yrange "-b/2", "b/2"
     .Zrange "0", "10" 
     .Create
    End With

'@ 创建 DoubleRidgeWaveGuide 被切除部分

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Brick
     .Reset 
     .Name "DoubleRidgeWaveGuidecutoff" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-s/2", "s/2" 
     .Yrange "d/2", "b/2"
     .Zrange "0", "10" 
     .Create
    End With

'@ 镜像 DoubleRidgeWaveGuide 切除部分

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Transform 
     .Reset 
     .Name "WaveGuide:DoubleRidgeWaveGuidecutoff" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "0", "-1", "0" 
     .MultipleObjects "True" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .Transform "Shape", "Mirror" 
    End With

'@ 开始减去 DoubleRidgeWaveGuide 切除部分部位1

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Solid.Subtract "WaveGuide:DRW1", "WaveGuide:DoubleRidgeWaveGuidecutoff"

'@ 开始减去 DoubleRidgeWaveGuide 切除部分部位2

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Solid.Subtract "WaveGuide:DRW1", "WaveGuide:DoubleRidgeWaveGuidecutoff_1"

'@ 添加过渡波导

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Brick
     .Reset 
     .Name "TW" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-ta/2", "ta/2" 
     .Yrange "-tb/2", "tb/2"
     .Zrange "0", "trh" 
     .Create
    End With

'@ 将脊波导与过渡波导相加

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Solid.Add "WaveGuide:DRW1", "WaveGuide:TW"

'@ 激活全局坐标系，准备变换

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
WCS.ActivateWCS "global"

'@ 将创建完成的脊波导进行镜像

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Transform 
     .Reset 
     .Name "WaveGuide:DRW1" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "0", "0", "-1" 
     .MultipleObjects "True" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .Transform "Shape", "Mirror" 
    End With

'@ 选取面1

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Pick.PickFaceFromId "WaveGuide:DRW1", "27"

'@ 添加端口1Add Port1

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Port 
        .Reset 
        .PortNumber "1" 
        .Label ""
        .Folder ""
        .NumberOfModes "1"
        .AdjustPolarization "False"
        .PolarizationAngle "0.0"
        .ReferencePlaneDistance "0"
        .TextSize "50"
        .TextMaxLimit "0"
        .Coordinates "Picks"
        .Orientation "positive"
        .PortOnBound "True"
        .ClipPickedPortToBound "False"
        .Xrange "-a/2", "a/2"
        .Xrange "-a/2", "a/2"
        .Yrange "wt/2+10", "wt/2+10"
        .XrangeAdd "0", "0"
        .XrangeAdd "0", "0"
        .ZrangeAdd "0", "0"
        .SingleEnded "False"
        .WaveguideMonitor "False"
        .Create 
    End With

'@ 选取面2

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Pick.PickFaceFromId "WaveGuide:DRW1_1", "27"

'@ 添加端口2Add Port2

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Port 
        .Reset 
        .PortNumber "2" 
        .Label ""
        .Folder ""
        .NumberOfModes "1"
        .AdjustPolarization "False"
        .PolarizationAngle "0.0"
        .ReferencePlaneDistance "0"
        .TextSize "50"
        .TextMaxLimit "0"
        .Coordinates "Picks"
        .Orientation "positive"
        .PortOnBound "True"
        .ClipPickedPortToBound "False"
        .Xrange "-a/2", "a/2"
        .Xrange "-a/2", "a/2"
        .Yrange "-(wt/2+10)", "-(wt/2+10)"
        .XrangeAdd "0", "0"
        .XrangeAdd "0", "0"
        .ZrangeAdd "0", "0"
        .SingleEnded "False"
        .WaveguideMonitor "False"
        .Create 
    End With

'@ 网格更新

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
With Mesh 
            .MeshType "Tetrahedral" 
            .SetCreator "High Frequency"
        End With 
        With MeshSettings 
            'MAX CELL - WAVELENGTH REFINEMENT 
            .Set "StepsPerWaveNear", "10" 
            .Set "StepsPerWaveFar", "5" 
            .Set "PhaseErrorNear", "0.02" 
            .Set "PhaseErrorFar", "0.02" 
            .Set "CellsPerWavelengthPolicy", "cellsperwavelength" 
            'MAX CELL - GEOMETRY REFINEMENT 
            .Set "StepsPerBoxNear", "6" 
            .Set "StepsPerBoxFar", "5" 
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

'@ 设置求解器

'[VERSION]2022.5|31.0.1|20220603[/VERSION]
Mesh.SetCreator "High Frequency" 

        With FDSolver
            .Reset 
            .SetMethod "Tetrahedral", "General purpose" 
            .OrderTet "Second" 
            .OrderSrf "First" 
            .Stimulation "All", "All" 
            .ResetExcitationList 
            .AutoNormImpedance "False" 
            .NormingImpedance "50" 
            .ModesOnly "False" 
            .ConsiderPortLossesTet "True" 
            .SetShieldAllPorts "False" 
            .AccuracyHex "1e-6" 
            .AccuracyTet "1e-5" 
            .AccuracySrf "1e-3" 
            .LimitIterations "False" 
            .MaxIterations "0" 
            .SetCalcBlockExcitationsInParallel "True", "True", "" 
            .StoreAllResults "False" 
            .StoreResultsInCache "False" 
            .UseHelmholtzEquation "True" 
            .LowFrequencyStabilization "True" 
            .Type "Direct" 
            .MeshAdaptionHex "False" 
            .MeshAdaptionTet "True" 
            .AcceleratedRestart "True" 
            .FreqDistAdaptMode "Distributed" 
            .NewIterativeSolver "True" 
            .TDCompatibleMaterials "False" 
            .ExtrudeOpenBC "False" 
            .SetOpenBCTypeHex "Default" 
            .SetOpenBCTypeTet "Default" 
            .AddMonitorSamples "True" 
            .CalcPowerLoss "True" 
            .CalcPowerLossPerComponent "False" 
            .StoreSolutionCoefficients "True" 
            .UseDoublePrecision "False" 
            .UseDoublePrecision_ML "True" 
            .MixedOrderSrf "False" 
            .MixedOrderTet "False" 
            .PreconditionerAccuracyIntEq "0.15" 
            .MLFMMAccuracy "Default" 
            .MinMLFMMBoxSize "0.3" 
            .UseCFIEForCPECIntEq "True" 
            .UseEnhancedCFIE2 "True" 
            .UseFastRCSSweepIntEq "false" 
            .UseSensitivityAnalysis "False" 
            .UseEnhancedNFSImprint "False" 
            .RemoveAllStopCriteria "Hex"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Hex", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Hex", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Hex", "False"
            .RemoveAllStopCriteria "Tet"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Tet", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Tet", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Tet", "False"
            .AddStopCriterion "All Probes", "0.05", "2", "Tet", "True"
            .RemoveAllStopCriteria "Srf"
            .AddStopCriterion "All S-Parameters", "0.01", "2", "Srf", "True"
            .AddStopCriterion "Reflection S-Parameters", "0.01", "2", "Srf", "False"
            .AddStopCriterion "Transmission S-Parameters", "0.01", "2", "Srf", "False"
            .SweepMinimumSamples "3" 
            .SetNumberOfResultDataSamples "5001" 
            .SetResultDataSamplingMode "Automatic" 
            .SweepWeightEvanescent "1.0" 
            .AccuracyROM "1e-4" 
            .AddSampleInterval "", "", "1", "Automatic", "True" 
            .AddSampleInterval "", "", "", "Automatic", "False" 
            .MPIParallelization "False"
            .UseDistributedComputing "False"
            .NetworkComputingStrategy "RunRemote"
            .NetworkComputingJobCount "3"
            .UseParallelization "True"
            .MaxCPUs "1024"
            .MaximumNumberOfCPUDevices "2"
        End With

        With IESolver
            .Reset 
            .UseFastFrequencySweep "True" 
            .UseIEGroundPlane "False" 
            .SetRealGroundMaterialName "" 
            .CalcFarFieldInRealGround "False" 
            .RealGroundModelType "Auto" 
            .PreconditionerType "Auto" 
            .ExtendThinWireModelByWireNubs "False" 
            .ExtraPreconditioning "False" 
        End With

        With IESolver
            .SetFMMFFCalcStopLevel "0" 
            .SetFMMFFCalcNumInterpPoints "6" 
            .UseFMMFarfieldCalc "True" 
            .SetCFIEAlpha "0.500000" 
            .LowFrequencyStabilization "False" 
            .LowFrequencyStabilizationML "True" 
            .Multilayer "False" 
            .SetiMoMACC_I "0.0001" 
            .SetiMoMACC_M "0.0001" 
            .DeembedExternalPorts "True" 
            .SetOpenBC_XY "True" 
            .OldRCSSweepDefintion "False" 
            .SetRCSOptimizationProperties "True", "100", "0.00001" 
            .SetAccuracySetting "Custom" 
            .CalculateSParaforFieldsources "True" 
            .ModeTrackingCMA "True" 
            .NumberOfModesCMA "3" 
            .StartFrequencyCMA "-1.0" 
            .SetAccuracySettingCMA "Default" 
            .FrequencySamplesCMA "0" 
            .SetMemSettingCMA "Auto" 
            .CalculateModalWeightingCoefficientsCMA "True" 
            .DetectThinDielectrics "True" 
        End With

