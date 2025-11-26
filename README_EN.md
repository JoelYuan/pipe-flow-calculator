# Pipe Flow Calculator

## Program Overview
This program calculates recommended flow velocity, volumetric flow rate, and mass flow rate based on pipe diameter, medium type, and pressure information, then generates an Excel format result report.

## Input File Format Requirements

### Supported File Formats
- Excel files (.xlsx)
- CSV files (.csv)

### Table Structure Requirements
The input table must contain the following three columns (fixed order):
1. **Pipe Diameter** - Pipe diameter value (supports numeric or mm-unit format)
2. **Medium** - Fluid medium name (must use program-supported medium types)
3. **Remarks** - Optional, can contain pressure information (supports MPa or bar units)

### Supported Medium Types
The program supports the following medium types. When filling, you must use one of these exact names:

#### Water and Aqueous Solutions
- 自来水 (Tap Water)
- 循环冷却水 (Circulating Cooling Water)
- 盐水 (Saline Water)

#### Steam System
- 饱和蒸汽 (Saturated Steam)
- 过热蒸汽 (Superheated Steam)
- 冷凝水回水 (Condensate Return)

#### Gas Medium
- 压缩空气 (Compressed Air)
- 天然气 (Natural Gas)
- 氧气 (Oxygen)

#### Special Fluids
- 液氨 (Liquid Ammonia)
- 硫酸 (Sulfuric Acid)
- 泥浆 (Slurry)

#### HVAC Specific
- 乙二醇溶液 (Ethylene Glycol Solution)
- 热水 (Hot Water)
- 高温烟气 (High Temperature Flue Gas)

### Usage Examples
```
Pipe Diameter,Medium,Remarks
100,自来水,Normal pipe
80,压缩空气,0.6MPa
50,饱和蒸汽,1.0MPa
```

## Pressure Information Extraction Rules
- Supports extracting pressure values from remarks (MPa or bar units)
- Format examples: "Pressure 0.8MPa" or "10bar compressed air"
- If no pressure information is provided, 0.5MPa is used by default for calculations

## How to Use
1. After running the program, a file selection dialog will appear
2. Select a properly formatted .xlsx or .csv input file
3. After processing, a save dialog will appear
4. Choose the save location and filename, the program will ensure .xlsx extension
5. View the generated result file containing calculated flows and design recommendations

## Output Result Description
The output Excel file contains the following columns:
- 管径(mm) (Pipe Diameter) - Input pipe diameter
- 介质 (Medium) - Input medium type
- 备注 (Remarks) - Input remarks
- 压力(MPa) (Pressure) - Extracted or default pressure value
- 推荐流速(m/s) (Recommended Velocity) - Calculated recommended velocity
- 体积流量(m³/h) (Volumetric Flow Rate) - Calculated volumetric flow
- 质量流量(t/h) (Mass Flow Rate) - Calculated mass flow
- 介质类别 (Medium Category) - Category of the medium
- 设计建议 (Design Recommendations) - Relevant design reference recommendations

## Notes
1. Medium names must exactly match the supported type list
2. Pipe diameter values should use numeric or mm-unit format
3. Ensure the input file has the correct column order and content
4. Result files will be saved in Excel format, which can be opened directly in Excel