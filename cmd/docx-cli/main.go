package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"os"

	"github.com/wyatsahar/docx"
)

func main() {
	input := flag.String("i", "", "输入模板路径 (Input template path)")
	output := flag.String("o", "output.docx", "输出路径 (Output path)")
	data := flag.String("d", "", "JSON 数据字符串或文件路径 (JSON data string or file path)")
	prefix := flag.String("p", "{{", "占位符前缀 (Placeholder prefix)")
	suffix := flag.String("s", "}}", "占位符后缀 (Placeholder suffix)")

	flag.Parse()

	if *input == "" || *data == "" {
		fmt.Println("用法: docx-cli -i template.docx -d '{\"name\":\"value\"}'")
		flag.PrintDefaults()
		return
	}

	// 解析数据
	var values map[string]string
	if err := json.Unmarshal([]byte(*data), &values); err != nil {
		// 尝试作为文件读取
		content, err := os.ReadFile(*data)
		if err != nil {
			fmt.Printf("错误: 无法解析 JSON 数据或读取数据文件: %v\n", err)
			return
		}
		if err := json.Unmarshal(content, &values); err != nil {
			fmt.Printf("错误: 数据文件 JSON 格式不正确: %v\n", err)
			return
		}
	}

	// 加载文档
	config := docx.Config{
		PlaceholderPrefix: *prefix,
		PlaceholderSuffix: *suffix,
	}
	doc, err := docx.LoadWithOptions(*input, config)
	if err != nil {
		fmt.Printf("错误: 加载文档失败: %v\n", err)
		return
	}
	defer doc.Close()

	// 替换
	doc.SetValue(values)

	// 保存
	if err := doc.SaveToFile(*output); err != nil {
		fmt.Printf("错误: 保存失败: %v\n", err)
		return
	}

	fmt.Printf("成功！已生成: %s\n", *output)
}
