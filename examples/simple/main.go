package main

import (
	"fmt"
	"os"

	"github.com/wyatsahar/docx"
)

func main() {
	// 1. 基本文件用法 (Basic file usage)
	doc, err := docx.Load("template.docx")
	if err != nil {
		fmt.Printf("无法加载模板: %v\n", err)
		return
	}
	defer doc.Close()

	doc.SetValue("name", "张三")
	doc.SaveToFile("output.docx")

	// 2. 高级用法：使用 io.Reader (Advanced usage: io.Reader)
	f, _ := os.Open("template.docx")
	defer f.Close()
	info, _ := f.Stat()

	// 自定义配置：使用 ${} 占位符 (Custom config: use ${})
	config := docx.Config{
		PlaceholderPrefix: "${",
		PlaceholderSuffix: "}",
	}

	doc2, err := docx.LoadFromReader(f, info.Size(), config)
	if err != nil {
		panic(err)
	}
	defer doc2.Close()

	doc2.SetValue("name", "李四")

	// 保存到 Buffer (Save to buffer)
	buf, _ := doc2.SaveToBuffer()
	fmt.Printf("生成文档成功，Buffer 大小: %d 字节\n", buf.Len())
}
