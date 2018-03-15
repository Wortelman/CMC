' freq_dom

Sub Main ()

'@ change solver type

ChangeSolverType "HF Frequency Domain" 


'@ define frequency range

Solver.FrequencyRange "0.1", "100" 


'@ define frequency domain solver parameters

Mesh.SetCreator "High Frequency" 

' msgBox "net voor Background"
With Background 
     .Reset 
     .Type ("normal")
     .XminSpace "0.0" 
     .XmaxSpace "0.0" 
     .YminSpace "0.0" 
     .YmaxSpace "0.0" 
     .ZminSpace "0.0" 
     .ZmaxSpace "0.0" 
     .ApplyInAllDirections "False" 
End With 


With FDSolver
     .Reset 
     .SetMethod "Tetrahedral", "General purpose" 
     .OrderTet "Second" 
     .OrderSrf "First" 
     .Stimulation "All", "1" 
     .ResetExcitationList 
     .AutoNormImpedance "False" 
     .NormingImpedance "50" 
     .ModesOnly "False" 
     .ConsiderPortLossesTet "True" 
     .SetShieldAllPorts "False" 
     .AccuracyHex "1e-6" 
     .AccuracyTet "1e-4" 
     .AccuracySrf "1e-3" 
     .LimitIterations "False" 
     .MaxIterations "0" 
     .SetCalculateExcitationsInParallel "True", "False", "" 
     .StoreAllResults "False" 
     .StoreResultsInCache "True" 
     .UseHelmholtzEquation "True" 
     .LowFrequencyStabilization "True" 
     .Type "Auto" 
     .MeshAdaptionHex "False" 
     .MeshAdaptionTet "False" 
     .AcceleratedRestart "False" 
     .FreqDistAdaptMode "Distributed" 
     .NewIterativeSolver "True" 
     .TDCompatibleMaterials "False" 
     .ExtrudeOpenBC "False" 
     .SetOpenBCTypeHex "Default" 
     .SetOpenBCTypeTet "Default" 
     .AddMonitorSamples "True" 
     .CalcStatBField "False" 
     .CalcPowerLoss "False" 
     .CalcPowerLossPerComponent "False" 
     .StoreSolutionCoefficients "True" 
     .UseDoublePrecision "False" 
     .UseDoublePrecision_ML "True" 
     .MixedOrderSrf "False" 
     .MixedOrderTet "True" 
     .PreconditionerAccuracyIntEq "0.15" 
     .MLFMMAccuracy "Default" 
     .MinMLFMMBoxSize "0.20" 
     .UseCFIEForCPECIntEq "true" 
     .UseFastRCSSweepIntEq "false" 
     .UseSensitivityAnalysis "False" 
     .SetStopSweepIfCriterionMet "True" 
     .SetSweepThreshold "S-Parameters", "0.01" 
     .UseSweepThreshold "S-Parameters", "True" 
     .SetSweepThreshold "Probes", "0.05" 
     .UseSweepThreshold "Probes", "True" 
     .SweepErrorChecks "2" 
     .SweepMinimumSamples "3" 
     .SweepConsiderAll "True" 
     .SweepConsiderReset 
     .SetNumberOfResultDataSamples "1001" 
     .SetResultDataSamplingMode "Automatic" 
     .SweepWeightEvanescent "1.0" 
     .AccuracyROM "1e-4" 
     .AddSampleInterval "0.1", "0.2", "30", "Logarithmic", "False"
     .AddSampleInterval "0.2", "100", "30", "Logarithmic", "False"
     .MPIParallelization "False"
     .UseDistributedComputing "False"
     .NetworkComputingStrategy "RunRemote"
     .NetworkComputingJobCount "3"
     .UseParallelization "True"
     .MaxCPUs "48"
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
     .SetAccuracySetting "Custom" 
     .CalculateSParaforFieldsources "True" 
     .NumberOfModesCMA "3" 
     .StartFrequencyCMA "-1.0" 
     .SetAccuracySettingCMA "Default" 
     .FrequencySamplesCMA "0" 
     .SetMemSettingCMA "Auto" 
End With


'@ s-parameter post processing: yz-matrices

PostProcess1D.ActivateOperation "yz-matrices", "TRUE" 


With Mesh 
     .MeshType "Tetrahedral" 
     .SetCreator "High Frequency"
End With 
With MeshSettings 
     .SetMeshType "Tet" 
     .Set "Version", 1%
     'MAX CELL - WAVELENGTH REFINEMENT 
     .Set "StepsPerWaveNear", "4" 
     .Set "StepsPerWaveFar", "4" 
     .Set "PhaseErrorNear", "0.02" 
     .Set "PhaseErrorFar", "0.02" 
     .Set "CellsPerWavelengthPolicy", "automatic" 
     'MAX CELL - GEOMETRY REFINEMENT 
     .Set "StepsPerBoxNear", "35" 
     .Set "StepsPerBoxFar", "30" 
     .Set "ModelBoxDescrNear", "maxedge" 
     .Set "ModelBoxDescrFar", "maxedge" 
     'MIN CELL 
     .Set "UseRatioLimit", "1" 
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
     .SetMeshType "Tet" 
     .Set "UseAnisoCurveRefinement", "1" 
     .Set "UseSameSrfAndVolMeshGradation", "1" 
     .Set "VolMeshGradation", "1.5" 
     .Set "VolMeshOptimization", "1" 
End With 
With MeshSettings 
     .SetMeshType "Unstr" 
     .Set "SmallFeatureSize", "0" 
     .Set "CoincidenceTolerance", "1e-006" 
     .Set "SelfIntersectionCheck", "1" 
     .Set "OptimizeForPlanarStructures", "0" 
End With 
With Mesh 
     .SetParallelMesherMode "Tet", "maximum" 
     .SetMaxParallelMesherThreads "Tet", "1" 
end With


' With Material 
'      .Reset 
'      .FrqType "all"
'      .Type "Normal"
'      .MaterialUnit "Frequency", "Hz"
'      .MaterialUnit "Geometry", "m"
'      .MaterialUnit "Time", "s"
'      .MaterialUnit "Temperature", "Kelvin"
'      .Epsilon "1"
'      .Mu "1"
'      .Sigma "0.0"
'      .TanD "0.0"
'      .TanDFreq "0.0"
'      .TanDGiven "False"
'      .TanDModel "ConstSigma"
'      .EnableUserConstTanDModelOrderEps "False"
'      .ConstTanDModelOrderEps "1"
'      .SetElParametricConductivity "False"
'      .ReferenceCoordSystem "Global"
'      .CoordSystemType "Cartesian"
'      .SigmaM "0"
'      .TanDM "0.0"
'      .TanDMFreq "0.0"
'      .TanDMGiven "False"
'      .TanDMModel "ConstSigma"
'      .EnableUserConstTanDModelOrderMu "False"
'      .ConstTanDModelOrderMu "1"
'      .SetMagParametricConductivity "False"
'      .DispModelEps  "None"
'      .DispModelMu "None"
'      .DispersiveFittingSchemeEps "Nth Order"
'      .MaximalOrderNthModelFitEps "10"
'      .ErrorLimitNthModelFitEps "0.1"
'      .UseOnlyDataInSimFreqRangeNthModelEps "False"
'      .DispersiveFittingSchemeMu "Nth Order"
'      .MaximalOrderNthModelFitMu "10"
'      .ErrorLimitNthModelFitMu "0.1"
'      .UseOnlyDataInSimFreqRangeNthModelMu "False"
'      .UseGeneralDispersionEps "False"
'      .UseGeneralDispersionMu "False"
'      .NonlinearMeasurementError "1e-1"
'      .NLAnisotropy "False"
'      .NLAStackingFactor "1"
'      .NLADirectionX "1"
'      .NLADirectionY "0"
'      .NLADirectionZ "0"
'      .Rho "0"
'      .ThermalType "Normal"
'      .ThermalConductivity "0"
'      .HeatCapacity "0"
'      .DynamicViscosity "0"
'      .Emissivity "0"
'      .MetabolicRate "0"
'      .BloodFlow "0"
'      .VoxelConvection "0"
'      .MechanicsType "Unused"
'      .Colour "0.6", "0.6", "0.6" 
'      .Wireframe "False" 
'      .Reflection "False" 
'      .Allowoutline "True" 
'      .Transparentoutline "False" 
'      .Transparency "0" 
'      .ChangeBackgroundMaterial
' End With 



End Sub
