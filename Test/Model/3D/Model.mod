'# MWS Version: Version 2022.4 - Apr 26 2022 - ACIS 31.0.1 -

'# length = mm
'# frequency = GHz
'# time = ns
'# frequency range: fmin = 8 fmax = 9
'# created = '[VERSION]2022.4|31.0.1|20220426[/VERSION]


'@ Background Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Background
.ResetBackground
.Type "PEC"
End With

'@ Unit Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Units
.Geometry "mm"
.Frequency "ghz"
.Time "ns"
End With

'@ Boundary Initial

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'@ �洢����

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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
          
            MakeSureParameterExists("wt", "0.20")
            SetParameterDescription  ( "wt", "wt" )
          
            MakeSureParameterExists("wr", "6.21")
            SetParameterDescription  ( "wr", "wr" )
          
            MakeSureParameterExists("t", "0.20")
            SetParameterDescription  ( "t", "t" )

'@ Add Material Sapphire

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'@ ����Բ��������ʯ��Ƭ

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'@ ѡȡԲ����Ƭ���ĵ�

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Pick.PickCenterpointFromId "Window:SapphireWindow", "3"

'@ �����ĵ��Ƶ�Բ����Ƭ����

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
WCS.AlignWCSWithSelected "Point"

'@ ����˫����������

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Brick
     .Reset 
     .Name "DRW" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-a/2", "a/2" 
     .Yrange "-b/2", "b/2"
     .Zrange "0", "10" 
     .Create
End With

'@ �������г�����

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Brick
     .Reset 
     .Name "cutoff" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-d/2", "d/2" 
     .Yrange "c/2", "b/2"
     .Zrange "0", "10" 
     .Create
End With

'@ �����г�����

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Transform 
     .Reset 
     .Name "WaveGuide:cutoff" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "0", "-1", "0" 
     .MultipleObjects "True" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .AutoDestination "True" 
     .Transform "Shape", "Mirror" 
End With

'@ ��ʼ��ȥ��λ1

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Solid.Subtract "WaveGuide:DRW", "WaveGuide:cutoff"

'@ ��ʼ��ȥ��λ2

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Solid.Subtract "WaveGuide:DRW", "WaveGuide:cutoff_1"

'@ ��ӹ��ɲ���

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Brick
     .Reset 
     .Name "TW" 
     .Component "WaveGuide" 
     .Material "Vacuum" 
     .Xrange "-a/2", "a/2" 
     .Yrange "-b/2*0.75", "b/2*0.75"
     .Zrange "0", "t" 
     .Create
End With

'@ ������������ɲ������

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Solid.Add "WaveGuide:DRW", "WaveGuide:TW"

'@ ����ȫ������ϵ��׼���任

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
WCS.ActivateWCS "global"

'@ ��������ɵļ��������о���

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
With Transform 
     .Reset 
     .Name "WaveGuide:DRW" 
     .Origin "Free" 
     .Center "0", "0", "0" 
     .PlaneNormal "0", "0", "-1" 
     .MultipleObjects "True" 
     .GroupObjects "False" 
     .Repetitions "1" 
     .MultipleSelection "False" 
     .Destination "" 
     .Material "" 
     .AutoDestination "True" 
     .Transform "Shape", "Mirror" 
End With

'@ ѡȡ��1

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Pick.PickFaceFromId "WaveGuide:DRW", "27"

'@ Add Port1

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'@ ѡȡ��2

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
Pick.PickFaceFromId "WaveGuide:DRW_1", "27"

'@ Add Port2

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

'@ �������

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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
        
        With Mesh 
            .Update 
        End With

'@ ���������

'[VERSION]2022.4|31.0.1|20220426[/VERSION]
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

