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

//设置标记
func ensureMacroCompleted(mark string) string {
	tmp := []rune(mark)
	star := tmp[:2]
	if string(star) != `${` && string(tmp[len(tmp)-1]) != `}` {
		return `${` + mark + `}`
	}
	return mark
}

func indexClonedVariables(xmlRow, mark string, n int) string {
	var bt bytes.Buffer
	reg := regexp.MustCompile(`\$\{(.*?)\}`)

	for i := 0; i < n; i++ {
		bt.WriteString(reg.ReplaceAllString(xmlRow, "${$1#"+strconv.Itoa(i)+`}`))
	}
	return bt.String()
}

//CloneRow 复制行 (标记 行数)
func (d *Docx) CloneRow(mark string, n int) {
	var tempxml []string
	var bt bytes.Buffer
	// //设置标记
	mark = ensureMacroCompleted(mark)
	//查询第一次出现的地方
	tagPos := strings.Index(d.MainPart, mark)

	//查找标记点开始行标签
	rowStart := findRowStart(d, tagPos)
	//查找标记点结束行标签
	rowEnd := findRowEnd(d, tagPos)

	xmlRow := getRow(d, rowStart, rowEnd)
	tempxml = append(tempxml, getRow(d, 0, rowStart))
	tempxml = append(tempxml, indexClonedVariables(xmlRow, mark, n))
	tempxml = append(tempxml, getRow(d, rowEnd, len(d.MainPart)))
	for _, v := range tempxml {
		bt.WriteString(v)
	}
	d.MainPart = bt.String()
}
