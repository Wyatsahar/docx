package docx

import (
	"os"
	"testing"
)

func TestDocx(t *testing.T) {
	// 准备：尝试加载测试文件
	if _, err := os.Stat("testdata/document_test.docx"); os.IsNotExist(err) {
		t.Skip("跳过测试：缺少 testdata/document_test.docx 文件")
	}

	// 1. 测试常规文件加载 (默认使用 {{}})
	doc, err := Load("./testdata/document_test.docx")
	if err != nil {
		t.Fatalf("加载失败: %v", err)
	}
	defer doc.Close()

	// 测试文本替换
	err = doc.SetValue("search", "替换成功")
	if err != nil {
		t.Errorf("设置文本值失败: %v", err)
	}

	// 2. 测试 io.Reader 加载
	f, err := os.Open("./testdata/document_test.docx")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	fi, _ := f.Stat()

	doc2, err := LoadFromReader(f, fi.Size(), DefaultConfig)
	if err != nil {
		t.Fatalf("LoadFromReader 失败: %v", err)
	}
	defer doc2.Close()

	// 3. 测试 WriteTo (内存操作)
	buf, err := doc.SaveToBuffer()
	if err != nil {
		t.Errorf("保存到 Buffer 失败: %v", err)
	}
	if buf.Len() == 0 {
		t.Errorf("Buffer 长度为 0")
	}
}
