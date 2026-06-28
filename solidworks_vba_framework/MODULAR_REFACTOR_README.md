# Modular SolidWorks VBA Macro Structure

This folder contains the refactored modular source layout. The old V2/V3 macros are kept for reference; the new workflow imports the shared modules, all flat part modules, the assembly layout module, and the master module into a clean SolidWorks macro project.

## Import Order

1. `shared\SW_Common.bas`
2. `shared\SW_PropertyHelpers.bas`
3. `shared\SW_FeatureHelpers.bas`
4. `shared\SW_AssemblyHelpers.bas`
5. `shared\SW_DrawingHelpers.bas`
6. every file under `parts\*.bas`
7. `assembly\SW_AutoFeeder_AssemblyLayout.bas`
8. `master\SW_AutoFeeder_Master.bas`

Run `Main` from `SW_AutoFeeder_Master.bas`.

## Ownership Rule

- Part geometry and machining features live in each part's own VBA file under `parts`.
- Part module file names use underscores, such as `CUS_001_CenterBasePlate.bas`, so SolidWorks VBA can import them cleanly.
- Keep module file names as valid VBA identifiers; avoid hyphens, spaces, Chinese punctuation, and overly long names.
- The master macro only calls `Build_xxx` functions and records generated paths.
- The assembly macro only places generated parts by coordinate.
- CSV output is only a run report, not the source of dimensions.

## Outputs

- `SW_Output\01_Parts`
- `SW_Output\02_Assemblies\ASM-000_AutoFeeder_Modular.SLDASM`
- `SW_Output\03_Drawings`
- `SW_Output\05_BOM\Modular_Build_BOM.csv`
- `SW_Output\99_Logs\macro_run_modular.log`

## Migration Map

- CUS-001 -> solidworks_vba_framework\parts\CUS_001_CenterBasePlate.bas
- CUS-002 -> solidworks_vba_framework\parts\CUS_002_FeedSideRail.bas
- CUS-003 -> solidworks_vba_framework\parts\CUS_003_TopGuideBlock.bas
- CUS-004 -> solidworks_vba_framework\parts\CUS_004_SmallGuideBlock.bas
- CUS-005 -> solidworks_vba_framework\parts\CUS_005_VerticalStopPlate.bas
- CUS-006 -> solidworks_vba_framework\parts\CUS_006_ClampBlock.bas
- CUS-007 -> solidworks_vba_framework\parts\CUS_007_LongSlideBase.bas
- CUS-008 -> solidworks_vba_framework\parts\CUS_008_UpperBridgePlate.bas
- CUS-009 -> solidworks_vba_framework\parts\CUS_009_ShortPressPlate.bas
- CUS-010 -> solidworks_vba_framework\parts\CUS_010_TallSidePlate.bas
- CUS-011 -> solidworks_vba_framework\parts\CUS_011_EntryTransitionBlock.bas
- CUS-012 -> solidworks_vba_framework\parts\CUS_012_CylinderMountPlate.bas
- CUS-013 -> solidworks_vba_framework\parts\CUS_013_SidePostPlate.bas
- CUS-014 -> solidworks_vba_framework\parts\CUS_014_VerticalNarrowPlate.bas
- CUS-015 -> solidworks_vba_framework\parts\CUS_015_SmallMountPlate.bas
- CUS-016 -> solidworks_vba_framework\parts\CUS_016_SmallClampPlate.bas
- CUS-017 -> solidworks_vba_framework\parts\CUS_017_PusherTallBlock.bas
- CUS-018 -> solidworks_vba_framework\parts\CUS_018_EndSupportBlock.bas
- CUS-019 -> solidworks_vba_framework\parts\CUS_019_FootClampBlock.bas
- CUS-020 -> solidworks_vba_framework\parts\CUS_020_SlideTopBlock.bas
- CUS-021 -> solidworks_vba_framework\parts\CUS_021_GuideRodClamp.bas
- CUS-022 -> solidworks_vba_framework\parts\CUS_022_GuideBearingBlock.bas
- CUS-023 -> solidworks_vba_framework\parts\CUS_023_LongGuideSupport.bas
- CUS-024 -> solidworks_vba_framework\parts\CUS_024_EndCylinderPlate.bas
- CUS-025 -> solidworks_vba_framework\parts\CUS_025_RodSeatBlock_D22.bas
- CUS-026 -> solidworks_vba_framework\parts\CUS_026_PositionPlate.bas
- CUS-027 -> solidworks_vba_framework\parts\CUS_027_LongLimitRail.bas
- CUS-028 -> solidworks_vba_framework\parts\CUS_028_SensorLightVertPlate.bas
- CUS-029 -> solidworks_vba_framework\parts\CUS_029_BackPanel.bas
- CUS-030 -> solidworks_vba_framework\parts\CUS_030_LongVerticalSupport.bas
- CUS-031 -> solidworks_vba_framework\parts\CUS_031_SensorAngleBlock.bas
- CUS-032 -> solidworks_vba_framework\parts\CUS_032_TubeClamp.bas
- CUS-033 -> solidworks_vba_framework\parts\CUS_033_SmallAdjuster.bas
- CUS-034 -> solidworks_vba_framework\parts\CUS_034_LinearFeederInsertPlate.bas
- CUS-035 -> solidworks_vba_framework\parts\CUS_035_StopperCylinderBracket.bas
- CUS-036 -> solidworks_vba_framework\parts\CUS_036_MainPusherStrokeRail.bas
- CUS-037 -> solidworks_vba_framework\parts\CUS_037_MainPusherRodSupport.bas
- FRM-001 -> solidworks_vba_framework\parts\FRM_001_MachineFrameEnvelope.bas
- PUR-001 -> solidworks_vba_framework\parts\PUR_001_LeftBowl.bas
- PUR-002 -> solidworks_vba_framework\parts\PUR_002_RightBowl.bas
- PUR-003 -> solidworks_vba_framework\parts\PUR_003_LeftLinearFeeder.bas
- PUR-004 -> solidworks_vba_framework\parts\PUR_004_RightLinearFeeder.bas
- PUR-005 -> solidworks_vba_framework\parts\PUR_005_MainCylinder.bas
- PUR-006 -> solidworks_vba_framework\parts\PUR_006_StopperCylinder.bas
- SM-001 -> solidworks_vba_framework\parts\SM_001_OptionalGuardPanel.bas
- STD-001 -> solidworks_vba_framework\parts\STD_001_GuideRodBearingSet.bas