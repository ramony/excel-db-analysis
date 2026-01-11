package main

import (
	"context"
	"database/sql"
	"encoding/csv"
	"fmt"
	"os"
	"strings"

	_ "github.com/mattn/go-sqlite3"
	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

// App 是 Wails 应用的核心结构体
type App struct {
	ctx context.Context
	db  *sql.DB
	// 存储最新查询结果
	queryResult struct {
		Columns []string
		Data    []map[string]interface{}
	}
}

// NewApp 创建 App 实例
func NewApp() *App {
	// 初始化 SQLite 数据库
	db, err := sql.Open("sqlite3", "./data.db")
	if err != nil {
		fmt.Printf("数据库连接失败: %v\n", err)
	}
	return &App{db: db}
}

// Startup 应用启动时执行
func (a *App) Startup(ctx context.Context) {
	a.ctx = ctx
}

// OpenExcel 选择并导入 Excel 文件（核心功能1）
// wails:export OpenExcel
func (a *App) OpenExcel() string {
	// 调用 Wails 原生文件选择对话框
	filePath, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{
		Title:                "选择 Excel 文件",
		Filters:              []runtime.FileFilter{{Pattern: "*.xlsx;*.xls", DisplayName: "Excel 文件"}},
		CanCreateDirectories: false,
	})
	if err != nil {
		return fmt.Sprintf("文件选择失败: %v", err)
	}
	if filePath == "" {
		return "未选择文件"
	}

	// 解析 Excel 文件
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Sprintf("Excel 解析失败: %v", err)
	}

	// 遍历所有 Sheet
	sheets := f.GetSheetList()
	for sheetIdx, sheetName := range sheets {
		tableName := fmt.Sprintf("sheet%d", sheetIdx+1)

		// 获取 Sheet 数据
		rows, err := f.GetRows(sheetName)
		if err != nil {
			return fmt.Sprintf("读取 Sheet %s 失败: %v", sheetName, err)
		}
		if len(rows) == 0 {
			continue
		}

		// 删除旧表
		_, err = a.db.Exec(fmt.Sprintf("DROP TABLE IF EXISTS %s", tableName))
		if err != nil {
			return fmt.Sprintf("删除表 %s 失败: %v", tableName, err)
		}

		// 创建新表（字段名 column1、column2...）
		colCount := len(rows[0])
		columns := make([]string, colCount)
		for i := 0; i < colCount; i++ {
			columns[i] = fmt.Sprintf("column%d", i+1)
		}

		createSQL := fmt.Sprintf(
			"CREATE TABLE %s (%s)",
			tableName,
			strings.Join(columns, " TEXT, ")+" TEXT",
		)
		_, err = a.db.Exec(createSQL)
		if err != nil {
			return fmt.Sprintf("创建表 %s 失败: %v", tableName, err)
		}

		// 插入数据（跳过表头）
		for rowIdx := 1; rowIdx < len(rows); rowIdx++ {
			row := rows[rowIdx]
			for len(row) < colCount {
				row = append(row, "")
			}

			placeholders := make([]string, colCount)
			values := make([]interface{}, colCount)
			for i := 0; i < colCount; i++ {
				placeholders[i] = "?"
				values[i] = row[i]
			}

			insertSQL := fmt.Sprintf(
				"INSERT INTO %s (%s) VALUES (%s)",
				tableName,
				strings.Join(columns, ", "),
				strings.Join(placeholders, ", "),
			)
			_, err = a.db.Exec(insertSQL, values...)
			if err != nil {
				return fmt.Sprintf("插入数据失败: %v", err)
			}
		}
	}

	return fmt.Sprintf("成功导入 %d 个 Sheet 到数据库", len(sheets))
}

// ExecuteSQL 执行 SQL 查询（核心功能2）
func (a *App) ExecuteSQL(sqlStr string) map[string]interface{} {
	result := make(map[string]interface{})
	sqlStr = strings.TrimSpace(sqlStr)
	if sqlStr == "" {
		result["error"] = "请输入 SQL 语句"
		return result
	}

	// 执行 SQL
	rows, err := a.db.Query(sqlStr)
	if err != nil {
		result["error"] = fmt.Sprintf("SQL 执行失败: %v", err)
		return result
	}
	defer rows.Close()

	// 获取列名
	columns, err := rows.Columns()
	if err != nil {
		result["error"] = fmt.Sprintf("获取列名失败: %v", err)
		return result
	}

	// 读取数据
	var data []map[string]interface{}
	values := make([]interface{}, len(columns))
	valuePtrs := make([]interface{}, len(columns))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	for rows.Next() {
		err := rows.Scan(valuePtrs...)
		if err != nil {
			result["error"] = fmt.Sprintf("读取数据失败: %v", err)
			return result
		}

		row := make(map[string]interface{})
		for i, col := range columns {
			val := values[i]
			if b, ok := val.([]byte); ok {
				row[col] = string(b)
			} else {
				row[col] = val
			}
		}
		data = append(data, row)
	}

	// 保存查询结果
	a.queryResult.Columns = columns
	a.queryResult.Data = data

	result["columns"] = columns
	result["data"] = data
	result["message"] = fmt.Sprintf("查询到 %d 条记录", len(data))
	return result
}

// SaveResult 保存查询结果到文件（核心功能3）
func (a *App) SaveResult() string {
	if len(a.queryResult.Data) == 0 {
		return "暂无查询结果可保存"
	}

	// 调用 Wails 保存文件对话框
	savePath, err := runtime.SaveFileDialog(a.ctx, runtime.SaveDialogOptions{
		Title:           "保存查询结果",
		DefaultFilename: "查询结果.csv",
		Filters:         []runtime.FileFilter{{Pattern: "*.csv", DisplayName: "CSV 文件"}},
	})
	if err != nil {
		return fmt.Sprintf("文件保存失败: %v", err)
	}
	if savePath == "" {
		return "取消保存"
	}

	// 写入 CSV 文件
	file, err := os.Create(savePath)
	if err != nil {
		return fmt.Sprintf("创建文件失败: %v", err)
	}
	defer file.Close()

	writer := csv.NewWriter(file)
	defer writer.Flush()

	// 写入表头
	if err := writer.Write(a.queryResult.Columns); err != nil {
		return fmt.Sprintf("写入表头失败: %v", err)
	}

	// 写入数据
	for _, row := range a.queryResult.Data {
		var record []string
		for _, col := range a.queryResult.Columns {
			val, ok := row[col]
			if !ok {
				record = append(record, "")
				continue
			}
			switch v := val.(type) {
			case string:
				record = append(record, v)
			case int, int64, float64:
				record = append(record, fmt.Sprintf("%v", v))
			default:
				record = append(record, fmt.Sprintf("%v", v))
			}
		}
		if err := writer.Write(record); err != nil {
			return fmt.Sprintf("写入数据失败: %v", err)
		}
	}

	return fmt.Sprintf("文件保存成功: %s", savePath)
}
