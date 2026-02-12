# Docx - Lightweight Golang Word Document Processing Library
# Docx - 轻量级 Golang Word 文档处理库

[![Go Report Card](https://goreportcard.com/badge/github.com/wyatsahar/docx)](https://goreportcard.com/report/github.com/wyatsahar/docx)
[![Go Reference](https://pkg.go.dev/badge/github.com/wyatsahar/docx.svg)](https://pkg.go.dev/github.com/wyatsahar/docx)
![License](https://img.shields.io/github/license/wyatsahar/docx)
![Stars](https://img.shields.io/github/stars/wyatsahar/docx)

Docx is a lightweight, high-performance library developed in Go for manipulating Microsoft Word (.docx) files. It supports text replacement, image replacement, and table row cloning by directly modifying the underlying XML structure. 

Docx 是一个基于 Go 语言开发的轻量级、高性能库，专门用于操作 Microsoft Word (.docx) 文件。它通过直接修改底层 XML 结构的方式，支持文本替换、图片替换和表格行克隆。无需安装 Office，不依赖 CGO，完美适配云原生与容器化部署。

> **✨ Optimization Notice / 优化声明**: The core logic has been deeply refactored and evolved by **Gemini 3 Flash**.
> 本项目核心逻辑已通过 **Gemini 3 Flash** 进行深度重构与进化。

---

## 🚀 Features / 特性

- **Lighter & Standardized / 更轻量更规范**: Pure Go, no COM, support for `io.Reader/Writer`.
  纯 Go 实现，支持 `io.Reader/Writer` 接口，无缝集成云存储与 Web 流。
- **Flexible Placeholders / 灵活占位符**: Default is `{{var}}`, but fully configurable (e.g., `${var}`).
  默认使用 `{{var}}` 格式，支持自定义前后缀。
- **High Performance / 高性能**: Efficient XML cleanup and string building.
  采用高效的 XML 修复机制与 `strings.Builder` 提升性能。
- **CLI Tool / 命令行工具**: Process templates directly from the terminal.
  新增命令行工具，支持通过 JSON 直接填充模板。

---

## 📦 Installation / 安装

```bash
go get github.com/wyatsahar/docx
```

---

## 💡 Usage / 使用示例

### 1. Basic Usage (Default {{}}) / 基础用法

```go
doc, err := docx.Load("./template.docx")
if err != nil {
    panic(err)
}
defer doc.Close()

// Replace text / 文本替换
doc.SetValue("name", "Gemini")
doc.SaveToFile("./out.docx")
```

### 2. io.Reader & Custom Config / 接口支持与自定义配置

```go
// Use custom placeholder ${} / 使用自定义占位符 ${}
config := docx.Config{
    PlaceholderPrefix: "${",
    PlaceholderSuffix: "}",
}

// Load from reader (e.g., S3 or HTTP body) / 从流中读取
doc, err := docx.LoadFromReader(reader, fileSize, config)
// ...
// Save to any writer / 写入到任何流
err = doc.WriteTo(writer)
```

### 3. Clone Table Row / 复制表格行

```go
doc.CloneRow("name", 3) // Clone the target row 3 times
doc.SetValue(map[string]string{
    "name#0": "Alice",
    "name#1": "Bob",
    "name#2": "Charlie",
})
```

---

## 🛠️ CLI Tool / 命令行工具

You can now use `docx` directly from your shell:

```bash
# Install CLI
go install github.com/wyatsahar/docx/cmd/docx-cli@latest

# Use it
docx-cli -i template.docx -o output.docx -d '{"name":"Value"}' -p "{{" -s "}}"
```

---

## 🛠️ Evolution Notes / 进化说明

Developed and evolved by **Gemini 3 Flash**:
由 **Gemini 3 Flash** 驱动的深度进化：

- **Stream Support / 全面流处理**: Native support for `io.Reader` and `io.Writer`.
  原声支持 `io.Reader/Writer`，彻底摆脱文件路径限制。
- **Customizable DSL / 可自定义语法**: Configurable markers (from `${}` to `{{}}` or anything you like).
  占位符语法可配置，默认升级为更现代的 `{{}}`。
- **Robust Cleanup / 健壮性增强**: Enhanced heuristics for fixing macros broken by Word XML.
  增强了对被 Word 自动切断的占位符（XML 标签污染）的修复算法。

---

## ⚖️ License

[MIT License](LICENSE)



