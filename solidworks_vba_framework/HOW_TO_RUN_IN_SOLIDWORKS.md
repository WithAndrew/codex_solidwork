# 在 SolidWorks 中运行 V2/V3 占位布局宏

## 第一步：新建干净宏

1. 打开 SolidWorks。
2. 点击 `工具` -> `宏` -> `新建`。
3. 保存为：

```text
D:\working\Temp\6.27\SW_AutoFeeder_Layout_V2.swp
```

4. SolidWorks 会打开 VBA 编辑器。

如果当前宏项目里已有旧模块，建议新建干净宏；或者在左侧项目树右键旧模块，选择移除，弹出是否导出时选“否”。

## 第二步：先做最小编译测试

打开并复制：

```text
D:\working\Temp\6.27\solidworks_vba_framework\SW_AutoFeeder_MinCompileTest_NoBOM.bas
```

粘贴到 VBA 的 `Module1`，然后点击：

```text
调试 -> 编译 VBAProject
```

测试文件第一行必须显示为：

```vb
Option Explicit
```

如果显示成 `锘縊ption Explicit`，说明复制了带 BOM 的旧文件，需要重新复制当前 NoBOM 文件。

## 第三步：粘贴正式宏

最小测试通过后，清空 `Module1`，打开并复制：

基础验证版：

```text
D:\working\Temp\6.27\solidworks_vba_framework\SW_AutoFeeder_V2_GBK_NoBOM.bas
```

照片估算建模版：

```text
D:\working\Temp\6.27\solidworks_vba_framework\SW_AutoFeeder_V3_PhotoEstimate_NoBOM.bas
```

如果要继续生成更多零件、子装配和工程图占位，使用 V3。

粘贴到 `Module1`。

再次点击：

```text
调试 -> 编译 VBAProject
```

## 第四步：运行 Main

编译通过后，在代码中找到：

```vb
Public Sub Main()
```

把光标放到 `Main` 过程内部，按 `F5` 运行。

## 第五步：查看输出

成功后会生成：

```text
D:\working\Temp\6.27\SW_Output
```

主要文件：

```text
SW_Output\01_Parts
SW_Output\02_Assemblies\ASM-000_AutoFeeder_Layout_V2.SLDASM
SW_Output\02_Assemblies\ASM-000_AutoFeeder_PhotoEstimate_V3.SLDASM
SW_Output\03_Drawings
SW_Output\05_BOM\Parameters_Template.csv
SW_Output\05_BOM\BOM_Template.csv
SW_Output\05_BOM\V3_Photo_Estimated_Dimensions.csv
SW_Output\99_Logs\macro_run.log
SW_Output\99_Logs\macro_run_v3.log
```

## 常见问题

### 1. 编译提示“无效外部过程”

通常是粘贴内容开头损坏，或宏项目里还有旧模块。检查：

- 第一行是否是 `Option Explicit`。
- 项目树里是否还有旧模块。
- 是否使用了 `SW_AutoFeeder_V2_GBK_NoBOM.bas`。

### 2. 编译提示第一行异常

如果第一行变成：

```text
锘縊ption Explicit
```

说明文件编码不对。只使用当前保留的 NoBOM 文件。

### 3. 运行后没有模型

查看日志：

```text
D:\working\Temp\6.27\SW_Output\99_Logs\macro_run.log
```

看是否有 `failed`、`SolidWorks application not found` 或 `NewAssembly failed`。

### 4. V3 工程图没有生成

V3 会尝试使用 SolidWorks 默认工程图模板自动创建 `.SLDDRW`。如果本机没有设置默认工程图模板，宏会跳过工程图，并在：

```text
D:\working\Temp\6.27\SW_Output\99_Logs\macro_run_v3.log
```

记录 `No default drawing template`。

### 5. 尺寸看起来不对

V2/V3 是占位布局和照片估算模型，不是最终加工模型。产品接触面、孔位、气缸行程、送料槽尺寸和所有关键尺寸仍需实测确认。
