package main

import (
	"bufio"
	"database/sql"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	_ "github.com/go-sql-driver/mysql"
	"github.com/tealeg/xlsx"
)

var mysqlHost, mysqlPort, mysqlUser, mysqlPassword, mysqlDb *string
var excelFilePath *string
var db *sql.DB

func main() {
	mysqlHost = flag.String("h", "localhost", "mysql host")
	mysqlPort = flag.String("P", "3306", "mysql port")
	mysqlDb = flag.String("d", "", "mysql database name")
	mysqlUser = flag.String("u", "", "mysql user name")
	mysqlPassword = flag.String("p", "", "mysql password")
	excelFilePath = flag.String("t", "", "export xlsx path")

	flag.Parse()

	if *mysqlHost == "" || *mysqlDb == "" || *excelFilePath == "" || *mysqlUser == "" {
		flag.PrintDefaults()
		return
	}

	excelAbsFilePath, err := filepath.Abs(*excelFilePath)
	if err != nil {
		fmt.Println(err)
		return
	}
	reader := bufio.NewReader(os.Stdin)

	fmt.Printf("Enter database password:\n")
	*mysqlPassword, _ = reader.ReadString('\n')
	*mysqlPassword = strings.TrimSpace(*mysqlPassword)

	//获取Sql语句
	var sqlStr string
	fmt.Printf("Please Input SQL:\n")
	for {
		tmpSql, _ := reader.ReadString('\n')
		sqlStr = sqlStr + tmpSql
		tmpSql = strings.TrimSpace(tmpSql)
		if tmpSql[len(tmpSql)-1] == ';' {
			break
		}
	}

	//获取sql连接
	dsn := *mysqlUser + ":" + *mysqlPassword + "@tcp(" + *mysqlHost + ":" + *mysqlPort + ")/" + *mysqlDb + "?charset=utf8"
	db, err := sql.Open("mysql", dsn)
	if err != nil {
		fmt.Println(err)
		return
	}

	//查询获取结果
	stmt, err := db.Prepare(sqlStr)
	if err != nil {
		fmt.Println(err)
		return
	}

	defer stmt.Close()

	result, err := stmt.Query()
	if err != nil {
		fmt.Println(err)
		return
	}

	//保存excel
	saveExcelByRows(excelAbsFilePath, result)
}

func saveExcelByRows(excelAbsFilePath string, rows *sql.Rows) error {
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("result")
	if err != nil {
		return err
	}

	columns, err := rows.Columns()
	if err != nil {
		return err
	}

	//写column数据
	columnRow := sheet.AddRow()
	columnLen := len(columns)
	for _, name := range columns {
		cell := columnRow.AddCell()
		cell.Value = name
	}

	scanArgs := make([]interface{}, columnLen)
	values := make([][]byte, columnLen)
	for i := range values {
		scanArgs[i] = &values[i]
	}
	for rows.Next() {
		rows.Scan(scanArgs...)
		row := sheet.AddRow()
		for _, v := range values {
			cell := row.AddCell()
			cell.Value = string(v)
		}
	}

	//保存文件
	err = file.Save(excelAbsFilePath)
	if err != nil {
		return err
	} else {
		return nil
	}

}
