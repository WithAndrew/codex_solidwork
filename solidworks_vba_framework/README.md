# SolidWorks VBA 宏说明

本目录保留当前已在 SolidWorks VBA 中调试编译通过的宏文件。

## 当前可用文件

- `SW_AutoFeeder_MinCompileTest_NoBOM.bas`：最小编译测试文件，用于确认 VBA 粘贴位置和编码正常。
- `SW_AutoFeeder_V2_GBK_NoBOM.bas`：正式 V2 单模块宏，已验证可以编译并生成 `SW_Output`。
- `SW_AutoFeeder_V3_PhotoEstimate_NoBOM.bas`：V3 照片估算建模宏，按 `picture/` 的 54 张照片和标注尺寸拆出更多独立零件、子装配和工程图占位。
- `templates/`：参数、BOM、零件登记模板。
- `templates/V3_Photo_Estimated_Dimensions.csv`：V3 照片估算尺寸登记表，单位 mm。
- `HOW_TO_RUN_IN_SOLIDWORKS.md`：SolidWorks 内部操作步骤。

## 编码要求

SolidWorks VBA 直接粘贴代码时，不要使用 UTF-8 with BOM 文件。BOM 会被粘贴成类似：

```text
锘縊ption Explicit
```

正确第一行必须是：

```vb
Option Explicit
```

当前保留的 `.bas` 文件均按可粘贴格式整理：无 `Attribute VB_Name`，无 UTF-8 BOM。

## 宏输出

运行 `SW_AutoFeeder_V2_GBK_NoBOM.bas` 或 `SW_AutoFeeder_V3_PhotoEstimate_NoBOM.bas` 后生成：

```text
D:\working\Temp\6.27\SW_Output
```

主要结果：

- `SW_Output\01_Parts`：占位零件和照片估算零件。
- `SW_Output\02_Assemblies`：总装布局和 V3 子装配。
- `SW_Output\03_Drawings`：V3 工程图占位文件。如果本机没有默认工程图模板，宏会在日志中记录并跳过。
- `SW_Output\05_BOM`：参数和 BOM 模板。
- `SW_Output\99_Logs\macro_run.log`：运行日志。

当前模型是照片估算布局模型，用于沟通、测绘和方案确认，不是加工图依据。所有尺寸都需要复核后才能投产。
