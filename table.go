package docx

import (
	"bytes"
	"regexp"
	"strconv"
	"strings"
)

func findRowStart(d *Docx, offset int) (rowStart int) {
	temp := []byte(d.MainPart)[:offset]
	rowStart = strings.LastIndex(string(temp), `<w:tr `)
	if rowStart == -1 {
		rowStart = strings.LastIndex(string(temp), `<w:tr>`)
	}
	return
}

func findRowEnd(d *Docx, offset int) (rowEnd int) {

	rowEnd = strings.Index(string([]byte(d.MainPart)[offset:]), `</w:tr>`)

	return rowEnd + offset + 7
}

func getRow(d *Docx, startPosition, endPosition int) (content string) {
	if endPosition == -1 {
		endPosition = len(d.MainPart)
	}
	contentTemp := []byte(d.MainPart)
	content = string(contentTemp[startPosition:endPosition])
	return
}

// 设置标记
func ensureMacroCompleted(d *Docx, mark string) string {
	if strings.HasPrefix(mark, d.Config.PlaceholderPrefix) && strings.HasSuffix(mark, d.Config.PlaceholderSuffix) {
		return mark
	}
	return d.Config.PlaceholderPrefix + mark + d.Config.PlaceholderSuffix
}

func indexClonedVariables(d *Docx, xmlRow, mark string, n int) string {
	var bt bytes.Buffer
	// 转义前缀和后缀用于正则
	prefix := regexp.QuoteMeta(d.Config.PlaceholderPrefix)
	suffix := regexp.QuoteMeta(d.Config.PlaceholderSuffix)
	reg := regexp.MustCompile(prefix + `(.*?)` + suffix)

	for i := 0; i < n; i++ {
		// 恢复为原始格式并加上索引
		bt.WriteString(reg.ReplaceAllString(xmlRow, d.Config.PlaceholderPrefix+"$1#"+strconv.Itoa(i)+d.Config.PlaceholderSuffix))
	}
	return bt.String()
}

// CloneRow 复制行 (标记 行数)
func (d *Docx) CloneRow(mark string, n int) {
	var tempxml []string
	var bt bytes.Buffer
	// //设置标记
	mark = ensureMacroCompleted(d, mark)
	//查询第一次出现的地方
	tagPos := strings.Index(d.MainPart, mark)

	//查找标记点开始行标签
	rowStart := findRowStart(d, tagPos)
	//查找标记点结束行标签
	rowEnd := findRowEnd(d, tagPos)

	xmlRow := getRow(d, rowStart, rowEnd)
	tempxml = append(tempxml, getRow(d, 0, rowStart))
	tempxml = append(tempxml, indexClonedVariables(d, xmlRow, mark, n))
	tempxml = append(tempxml, getRow(d, rowEnd, len(d.MainPart)))
	for _, v := range tempxml {
		bt.WriteString(v)
	}
	d.MainPart = bt.String()
}
