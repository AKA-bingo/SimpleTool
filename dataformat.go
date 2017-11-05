package main

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
	"strconv"
)

func main() {
	args := os.Args //获取参数输入
	if args == nil || len(args) < 2 {
		return
	}
	jsonFile := args[1]  //json文件
	execlFile := args[2] //要生成的xlsx文件
	if !strings.Contains(execlFile, ".xlsx") {
		execlFile = execlFile + ".xlsx"
	}
	jsondata, err := readFile(jsonFile) //读取json文件
	if err != nil {
		return
	}
	//fmt.Println(jsondata)
	createExec(execlFile, jsondata) //生成xlsx文件
}

//读取文件
func readFile(filename string) (interface{}, error) {
	bytes, err := ioutil.ReadFile(filename)
	if err != nil {
		fmt.Println("ReadFile: ", err.Error())
		return nil, err
	}
	var data interface{}
	if err := json.Unmarshal(bytes, &data); err != nil {
		fmt.Println("Unmarshal: ", err.Error())
		return nil, err
	}
	return data, nil
}

func createExec(filename string, jsonData interface{}) {
	file := xlsx.NewFile()
	startcol := 0
	sheet, err := file.AddSheet("Sheet1")
	title := make(map[string]int)
	if err != nil {
		fmt.Println(err.Error())
	}
	for i, data := range jsonData.([]interface{}) {
		for col, value := range data.(map[string]interface{}) {
			switch t := value.(type) {
			case float64:
				value = strconv.FormatFloat(float64(value.(float64)), 'f', -1, 64)
			case string:
				value = t
			}
			if _, ok := title[col]; ok {
				cell := sheet.Cell(i+1, title[col])
				cell.Value = value.(string)
			} else {
				title[col] = startcol
				startcol++
				cell := sheet.Cell(i+1, title[col])
				cell.Value = value.(string)
			}
		}
	}
	for value, key := range title {
		cell := sheet.Cell(0, key)
		cell.Value = value
	}
	err = file.Save(filename)
	if err != nil {
		fmt.Println(err.Error())
	}
}
