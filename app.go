package main

import (
	"context"
	"database/sql"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	_ "github.com/mattn/go-sqlite3"
	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

// App 核心结构体（移除 fullResult 缓存）
type App struct {
	ctx             context.Context
	db              *sql.DB
	currentPage     int    // 当前页码
	currentPageSize int    // 当前页大小
	currentSQL      string // 保存当前执行的 SQL（用于分页）
}

// NewApp 创建 App 实例（完善数据库初始化）
func NewApp() *App {
	// 初始化 SQLite 数据库
	db, err := sql.Open("sqlite3", "./data.db")
	if err != nil {
		fmt.Printf("数据库连接失败: %v\n", err)
		// 创建数据库目录（避免路径不存在）
		os.MkdirAll(filepath.Dir("./data.db"), 0755)
		db, err = sql.Open("sqlite3", "./data.db")
		if err != nil {
			fmt.Printf("数据库重试连接失败: %v\n", err)
			return &App{db: nil}
		}
	}

	// 验证数据库连接
	if err := db.Ping(); err != nil {
		fmt.Printf("数据库 Ping 失败: %v\n", err)
		return &App{db: nil}
	}

	return &App{
		db:              db,
		currentPage:     1,
		currentPageSize: 20,
		currentSQL:      "",
	}
}

// Startup 应用启动时执行
func (a *App) Startup(ctx context.Context) {
	a.ctx = ctx
}

// OpenExcel 导入 Excel 文件（原有逻辑保留）
// wails:export OpenExcel
func (a *App) OpenExcel() string {
	if a.db == nil {
		return "错误：数据库连接未初始化，请重启应用！"
	}

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

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Sprintf("Excel 解析失败: %v", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	successCount := 0
	for sheetIdx, sheetName := range sheets {
		tableName := fmt.Sprintf("sheet%d", sheetIdx+1)
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

		// 创建新表
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

		// 批量插入数据
		tx, err := a.db.Begin()
		if err != nil {
			return fmt.Sprintf("开启事务失败: %v", err)
		}

		insertSQL := fmt.Sprintf(
			"INSERT INTO %s (%s) VALUES (%s)",
			tableName,
			strings.Join(columns, ", "),
			strings.Repeat("?,", colCount)[:len(strings.Repeat("?,", colCount))-1],
		)
		stmt, err := tx.Prepare(insertSQL)
		if err != nil {
			tx.Rollback()
			return fmt.Sprintf("预编译插入语句失败: %v", err)
		}
		defer stmt.Close()

		for rowIdx := 1; rowIdx < len(rows); rowIdx++ {
			row := rows[rowIdx]
			for len(row) < colCount {
				row = append(row, "")
			}

			values := make([]interface{}, colCount)
			for i := 0; i < colCount; i++ {
				values[i] = row[i]
			}

			_, err := stmt.Exec(values...)
			if err != nil {
				tx.Rollback()
				return fmt.Sprintf("插入第 %d 行数据失败: %v", rowIdx, err)
			}
		}

		if err := tx.Commit(); err != nil {
			return fmt.Sprintf("提交事务失败: %v", err)
		}
		successCount++
	}

	return fmt.Sprintf("成功导入 %d 个 Sheet 到数据库（共 %d 个 Sheet）", successCount, len(sheets))
}

// ExecuteSQLWithPage 执行分页 SQL 查询（保留分页功能）
// wails:export ExecuteSQLWithPage
func (a *App) ExecuteSQLWithPage(sqlStr string, pageNum int, pageSize int) map[string]interface{} {
	result := make(map[string]interface{})

	if a.db == nil {
		result["error"] = "错误：数据库连接未初始化，请重启应用！"
		return result
	}

	sqlStr = strings.TrimSpace(sqlStr)
	if sqlStr == "" {
		result["error"] = "请输入 SQL 语句"
		return result
	}

	// 保存当前执行的 SQL（用于分页跳转）
	a.currentSQL = sqlStr
	a.currentPage = pageNum
	a.currentPageSize = pageSize

	// 执行原始 SQL 获取全量数据（用于计算总数和内存分页）
	fullRows, err := a.db.Query(sqlStr)
	if err != nil {
		result["error"] = fmt.Sprintf("SQL 执行失败: %v", err)
		return result
	}
	defer fullRows.Close()

	// 获取列名
	columns, err := fullRows.Columns()
	if err != nil {
		result["error"] = fmt.Sprintf("获取列名失败: %v", err)
		return result
	}

	// 解析全量数据
	var fullData []map[string]interface{}
	values := make([]interface{}, len(columns))
	valuePtrs := make([]interface{}, len(columns))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	for fullRows.Next() {
		err := fullRows.Scan(valuePtrs...)
		if err != nil {
			result["error"] = fmt.Sprintf("读取数据失败: %v", err)
			return result
		}

		row := make(map[string]interface{})
		for i, col := range columns {
			val := values[i]
			if b, ok := val.([]byte); ok {
				row[col] = string(b)
			} else if val == nil {
				row[col] = ""
			} else {
				row[col] = val
			}
		}
		fullData = append(fullData, row)
	}

	if err = fullRows.Err(); err != nil {
		result["error"] = fmt.Sprintf("遍历数据失败: %v", err)
		return result
	}

	// 计算分页参数
	total := len(fullData)
	totalPages := (total + pageSize - 1) / pageSize

	// 内存分页
	start := (pageNum - 1) * pageSize
	end := start + pageSize
	if end > total {
		end = total
	}
	var pageData []map[string]interface{}
	if start < total {
		pageData = fullData[start:end]
	}

	// 返回分页结果
	result["columns"] = columns
	result["data"] = pageData
	result["total"] = total
	result["totalPages"] = totalPages
	result["currentPage"] = pageNum
	result["pageSize"] = pageSize
	result["message"] = fmt.Sprintf("查询到 %d 条记录，当前第 %d 页（共 %d 页）", total, pageNum, totalPages)
	return result
}

// ExportExcelBySQL 根据 SQL 实时查询并导出 Excel（核心重构）
// wails:export ExportExcelBySQL
func (a *App) ExportExcelBySQL(sqlStr string) string {
	// 1. 前置检查
	if a.db == nil {
		return "错误：数据库连接未初始化，请重启应用！"
	}

	sqlStr = strings.TrimSpace(sqlStr)
	if sqlStr == "" {
		return "错误：SQL 语句不能为空！"
	}

	// 2. 实时执行 SQL 获取全量数据（无分页）
	fullRows, err := a.db.Query(sqlStr)
	if err != nil {
		return fmt.Sprintf("SQL 执行失败: %v", err)
	}
	defer fullRows.Close()

	// 3. 获取列名
	columns, err := fullRows.Columns()
	if err != nil {
		return fmt.Sprintf("获取列名失败: %v", err)
	}

	// 4. 解析全量数据
	var fullData []map[string]interface{}
	values := make([]interface{}, len(columns))
	valuePtrs := make([]interface{}, len(columns))
	for i := range values {
		valuePtrs[i] = &values[i]
	}

	for fullRows.Next() {
		err := fullRows.Scan(valuePtrs...)
		if err != nil {
			return fmt.Sprintf("读取数据失败: %v", err)
		}

		row := make(map[string]interface{})
		for i, col := range columns {
			val := values[i]
			// 处理特殊类型，避免 Excel 写入空值
			if b, ok := val.([]byte); ok {
				row[col] = string(b)
			} else if val == nil {
				row[col] = ""
			} else {
				row[col] = val
			}
		}
		fullData = append(fullData, row)
	}

	// 检查遍历错误
	if err = fullRows.Err(); err != nil {
		return fmt.Sprintf("遍历数据失败: %v", err)
	}

	// 5. 检查数据是否为空
	if len(fullData) == 0 {
		return "导出失败：SQL 查询结果为空！"
	}

	fmt.Printf("[DEBUG] 共读取到 %d 行数据\n", len(fullData))

	// 6. 选择保存路径
	savePath, err := runtime.SaveFileDialog(a.ctx, runtime.SaveDialogOptions{
		Title:           "导出 Excel 文件",
		DefaultFilename: "查询结果.xlsx",
		Filters:         []runtime.FileFilter{{Pattern: "*.xlsx", DisplayName: "Excel 文件"}},
	})
	if err != nil {
		return fmt.Sprintf("文件保存失败: %v", err)
	}
	if savePath == "" {
		return "取消导出"
	}

	// 7. 生成 Excel 文件
	f := excelize.NewFile()
	defer f.Close()
	sheetName := "Sheet1"

	// 写入表头
	for colIdx, colName := range columns {
		cell := fmt.Sprintf("%c1", 'A'+(colIdx))
		f.SetCellValue(sheetName, cell, colName)
	}

	// 写入全量数据
	for rowIdx, rowData := range fullData {
		for colIdx, colName := range columns {
			cell := fmt.Sprintf("%c%d", 'A'+(colIdx), rowIdx+2)
			f.SetCellValue(sheetName, cell, rowData[colName])
		}
	}

	// 8. 保存文件
	if err := f.SaveAs(savePath); err != nil {
		return fmt.Sprintf("导出 Excel 失败: %v", err)
	}

	return fmt.Sprintf("Excel 导出成功: %s（共 %d 条数据）", savePath, len(fullData))
}

// GetCurrentSQL 获取当前执行的 SQL（用于前端导出）
// wails:export GetCurrentSQL
func (a *App) GetCurrentSQL() string {
	return a.currentSQL
}
