package main

import (
	"bufio"
	"fmt"
	"log"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/sqweek/dialog"
	"github.com/xuri/excelize/v2"
)

// 介质流速数据库 - 扁平化设计
var velocityDB = map[string]struct {
	min, max       float64
	category       string
	recommendation string
}{
	// 水及水溶液
	"自来水":   {1.0, 1.5, "水及水溶液", "防噪音要求≤1.2m/s"},
	"循环冷却水": {2.0, 2.5, "水及水溶液", "防腐蚀需≤2.2m/s"},
	"盐水":    {1.2, 1.8, "水及水溶液", "制冷系统/化工流程，需考虑沸点升高效应"},

	// 蒸汽系统
	"饱和蒸汽":  {20.0, 30.0, "蒸汽系统", "避免冷凝水携带（≤25m/s）"},
	"过热蒸汽":  {35.0, 50.0, "蒸汽系统", "管道振动控制"},
	"冷凝水回水": {0.5, 1.2, "蒸汽系统", "防气蚀设计"},

	// 气体介质
	"压缩空气": {10.0, 15.0, "气体介质", "气动工具管网，需设油水分离器"},
	"天然气":  {8.0, 12.0, "气体介质", "城市输配管网，含硫气体需降速20%"},
	"氧气":   {5.0, 8.0, "气体介质", "钢铁冶炼供气，禁油设计+流速下限控制"},

	// 特殊流体
	"液氨": {0.8, 1.5, "特殊流体", "保冷管道+防震支架"},
	"硫酸": {0.6, 1.2, "特殊流体", "衬塑管道+低流速防结晶"},
	"泥浆": {1.5, 2.0, "特殊流体", "流速需＞沉降临界值"},

	// 暖通专用
	"乙二醇溶液": {1.0, 2.5, "暖通专用", "ASHRAE标准"},
	"热水":    {0.3, 0.5, "暖通专用", "防气阻设计"},
	"高温烟气":  {8.0, 12.0, "暖通专用", "耐火材料内衬"},
}

func readCSVFile(filename string) ([][]string, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	var rows [][]string
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		row := strings.Split(scanner.Text(), ",")
		// 处理每列的数据，去除首尾空格
		for i := range row {
			row[i] = strings.TrimSpace(row[i])
		}
		rows = append(rows, row)
	}

	return rows, scanner.Err()
}

func main() {
	// 显示文件选择对话框
	filename, err := dialog.File().
		Title("选择输入文件").
		Filter("Excel文件", "xlsx").
		Filter("CSV文件", "csv").
		Load()

	if err != nil || filename == "" {
		log.Fatal("未选择文件或操作被取消")
	}

	var rows [][]string
	var f *excelize.File

	// 根据文件扩展名选择读取方式
	ext := strings.ToLower(filepath.Ext(filename))
	switch ext {
	case ".xlsx":
		// 读取Excel文件
		f, err = excelize.OpenFile(filename)
		if err != nil {
			log.Fatal("Error opening Excel file:", err)
		}
		defer f.Close()

		// 获取第一个工作表名称
		sheetName := f.GetSheetName(0)
		rows, err = f.GetRows(sheetName)
		if err != nil {
			log.Fatal("Error reading Excel rows:", err)
		}
	case ".csv":
		rows, err = readCSVFile(filename)
		if err != nil {
			log.Fatal("Error reading CSV file:", err)
		}
		// 为CSV文件创建Excel工作簿用于输出
		f = excelize.NewFile()
		defer f.Close()
	default:
		log.Fatal("不支持的文件格式，请选择.xlsx或.csv文件")
	}

	if len(rows) == 0 {
		log.Fatal("Excel file is empty")
	}

	// 检查是否有标题行，如果有则跳过
	startRow := 0
	if len(rows) > 0 && len(rows[0]) >= 3 {
		// 检查第一行是否包含预期的列标题
		firstRow := rows[0]
		if len(firstRow) >= 3 &&
			(strings.Contains(strings.ToLower(firstRow[0]), "管径") ||
				strings.Contains(strings.ToLower(firstRow[0]), "diameter")) {
			startRow = 1 // 有标题行，从第二行开始处理
		}
	}

	// 准备输出工作表
	outputSheet := "流量设计结果"
	f.NewSheet(outputSheet)

	// 写入标题
	headers := []string{
		"管径(mm)", "介质", "备注", "压力(MPa)", "推荐流速(m/s)",
		"体积流量(m³/h)", "质量流量(t/h)", "介质类别", "设计建议",
	}
	for i, header := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(outputSheet, cell, header)
	}

	rowNum := 2 // 从第二行开始写入数据
	for i := startRow; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 3 {
			continue
		}

		pipeDiameterStr := strings.TrimSpace(row[0])
		medium := strings.TrimSpace(row[1])
		remarks := strings.TrimSpace(row[2])

		pipeDiameter, err := parseFloat(pipeDiameterStr)
		if err != nil {
			continue
		}

		// 从备注中提取压力信息
		pressure := extractPressure(remarks)

		// 获取推荐流速
		velocity, category, recommendation := getRecommendedVelocity(medium)
		volumeFlowRate := calculateVolumeFlowRate(pipeDiameter, velocity)

		// 计算质量流量
		massFlowRate := 0.0
		if strings.Contains(medium, "蒸汽") {
			density := calculateSteamDensity(pressure)
			massFlowRate = volumeFlowRate * density / 1000 // 转换为t/h
		} else {
			// 对于其他介质，使用近似密度
			density := getApproximateDensity(medium)
			massFlowRate = volumeFlowRate * density / 1000 // 转换为t/h
		}

		// 写入结果到Excel
		f.SetCellValue(outputSheet, fmt.Sprintf("A%d", rowNum), pipeDiameter)
		f.SetCellValue(outputSheet, fmt.Sprintf("B%d", rowNum), medium)
		f.SetCellValue(outputSheet, fmt.Sprintf("C%d", rowNum), remarks)
		f.SetCellValue(outputSheet, fmt.Sprintf("D%d", rowNum), fmt.Sprintf("%.3f", pressure))
		f.SetCellValue(outputSheet, fmt.Sprintf("E%d", rowNum), fmt.Sprintf("%.2f", velocity))
		f.SetCellValue(outputSheet, fmt.Sprintf("F%d", rowNum), fmt.Sprintf("%.2f", volumeFlowRate))
		f.SetCellValue(outputSheet, fmt.Sprintf("G%d", rowNum), fmt.Sprintf("%.2f", massFlowRate))
		f.SetCellValue(outputSheet, fmt.Sprintf("H%d", rowNum), category)
		f.SetCellValue(outputSheet, fmt.Sprintf("I%d", rowNum), recommendation)

		rowNum++
	}

	// 显示保存文件对话框
	outputFilename, err := dialog.File().
		Title("保存结果文件").
		Filter("Excel文件", "xlsx").
		Save()

	if err != nil || outputFilename == "" {
		log.Fatal("未指定保存文件或操作被取消")
	}

	// 确保文件扩展名是.xlsx
	if !strings.HasSuffix(strings.ToLower(outputFilename), ".xlsx") {
		outputFilename += ".xlsx"
	}

	if err := f.SaveAs(outputFilename); err != nil {
		log.Fatal("Error saving Excel file:", err)
	}

	fmt.Printf("处理完成，结果已保存到: %s\n", outputFilename)
}

// parseFloat 解析浮点数，处理可能的格式问题
func parseFloat(s string) (float64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, fmt.Errorf("empty string")
	}

	// 移除可能的单位
	s = strings.ReplaceAll(s, "mm", "")
	s = strings.ReplaceAll(s, "MM", "")
	s = strings.TrimSpace(s)

	return strconv.ParseFloat(s, 64)
}

// extractPressure 从备注中提取压力值
func extractPressure(remarks string) float64 {
	// 查找压力关键词
	lowerRemarks := strings.ToLower(remarks)

	// 查找MPa
	if idx := strings.Index(lowerRemarks, "mpa"); idx > 0 {
		// 向前查找数字
		start := idx
		for start > 0 && (lowerRemarks[start-1] >= '0' && lowerRemarks[start-1] <= '9' || lowerRemarks[start-1] == '.' || lowerRemarks[start-1] == '-') {
			start--
		}
		if start < idx {
			if val, err := strconv.ParseFloat(remarks[start:idx], 64); err == nil {
				return val
			}
		}
	}

	// 查找bar
	if idx := strings.Index(lowerRemarks, "bar"); idx > 0 {
		start := idx
		for start > 0 && (lowerRemarks[start-1] >= '0' && lowerRemarks[start-1] <= '9' || lowerRemarks[start-1] == '.' || lowerRemarks[start-1] == '-') {
			start--
		}
		if start < idx {
			if val, err := strconv.ParseFloat(remarks[start:idx], 64); err == nil {
				return val * 0.1 // bar转MPa
			}
		}
	}

	// 默认值
	return 0.5 // MPa
}

// calculateSteamDensity 根据压力计算蒸汽密度
func calculateSteamDensity(pressureMPa float64) float64 {
	if pressureMPa < 0.32 {
		return 5.2353*pressureMPa + 0.0816
	} else if pressureMPa < 1.00 {
		return 5.0221*pressureMPa + 0.1517
	} else {
		return 4.9283*pressureMPa + 0.2173
	}
}

// getApproximateDensity 获取近似密度
func getApproximateDensity(medium string) float64 {
	switch {
	case strings.Contains(medium, "水"):
		return 1000 // kg/m³
	case strings.Contains(medium, "蒸汽"):
		return 1 // kg/m³ (近似值，实际会根据压力变化)
	case strings.Contains(medium, "空气"):
		return 1.2 // kg/m³
	case strings.Contains(medium, "氧气"):
		return 1.43 // kg/m³
	case strings.Contains(medium, "天然气"):
		return 0.7 // kg/m³
	case strings.Contains(medium, "氨"):
		return 0.77 // kg/m³
	case strings.Contains(medium, "硫酸"):
		return 1840 // kg/m³
	default:
		return 1000 // 默认水的密度
	}
}

// getRecommendedVelocity 根据介质获取推荐流速
func getRecommendedVelocity(medium string) (float64, string, string) {
	medium = strings.TrimSpace(medium)

	// 精确匹配
	if v, exists := velocityDB[medium]; exists {
		return (v.min + v.max) / 2, v.category, v.recommendation
	}

	// 模糊匹配
	for key, v := range velocityDB {
		if strings.Contains(medium, key) || strings.Contains(key, medium) {
			return (v.min + v.max) / 2, v.category, v.recommendation
		}
	}

	return 1.5, "未知", "请根据实际情况确定流速"
}

// calculateVolumeFlowRate 计算体积流量 (m³/h)
func calculateVolumeFlowRate(diameterMM, velocity float64) float64 {
	diameterM := diameterMM / 1000.0
	area := math.Pi * (diameterM / 2) * (diameterM / 2)
	flowMS := area * velocity
	return flowMS * 3600
}
