package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io/ioutil"
	"path/filepath"
	"regexp"
	"strings"
	"unsafe"
)

//Docx 文档
type Docx struct {
	MainPart     string
	SettingsPart string
	ContentTypes string
	Headers      map[int]string
	Footers      map[int]string
	Relations    map[string]string
	// NewImages    map[string]string
}

//ZipData Contains functions to work with data from a zip file
type ZipData interface {
	files() []*zip.File
	close() error
}

//ZipBuffer zip buffer
type ZipBuffer struct {
	rc *zip.ReadCloser
}

func (b ZipBuffer) files() []*zip.File {
	return b.rc.File
}

//Close 关闭
func (b ZipBuffer) Close() error {
	return b.rc.Close()
}

//LoadInit 初始化Docx
func LoadInit(path string) (*Docx, *ZipBuffer) {
	// func LoadInit(path string) {
	//打开zip文件
	rc, _ := zip.OpenReader(path)

	b := ZipBuffer{rc}
	Relations := make(map[string]string)
	MainPart := b.getTempDocumentMainPart(Relations)
	Headers := b.getTempDocumentHeaders(Relations)
	Footers := b.getTempDocumentFooters(Relations)
	SettingsPart := b.getTempDocumentSettingsPart(Relations)
	ContentTypes := b.getTempDocumentContentTypes(Relations)

	return &Docx{
		MainPart:     MainPart,
		Headers:      Headers,
		Footers:      Footers,
		Relations:    Relations,
		SettingsPart: SettingsPart,
		ContentTypes: ContentTypes,
	}, &b
}

//SetValue 替换文本
func (d *Docx) SetValue(search, replace string, limit int) error {
	encodeSearch, err := encode(search)
	if err != nil {
		return err
	}
	encodeReplace, err := encode(replace)
	if err != nil {
		return err
	}
	d.setValueForPart(encodeSearch, encodeReplace, limit)
	return nil
}

func encode(s string) (string, error) {
	var b bytes.Buffer
	enc := xml.NewEncoder(bufio.NewWriter(&b))
	if err := enc.Encode(s); err != nil {
		return s, err
	}
	output := strings.Replace(b.String(), "<string>", "", 1) // remove string tag
	output = strings.Replace(output, "</string>", "", 1)
	output = strings.Replace(output, "&#xD;&#xA;", "<w:br/>", -1) // \r\n => newline
	return output, nil
}

func (d *Docx) setValueForPart(search, replace string, limit int) {
	if limit <= 0 {
		limit = -1
	}
	strings.Replace(d.MainPart, search, replace, limit)

	for _, header := range d.Headers {
		strings.Replace(header, search, replace, limit)
	}
	for _, footer := range d.Footers {
		strings.Replace(footer, search, replace, limit)
	}
}

func (b *ZipBuffer) readPartWithRels(fileName string) string {
	return b.getFromName(getRelationsName(fileName))
}

func (b *ZipBuffer) getTempDocumentFooters(relations map[string]string) map[int]string {
	footers := make(map[int]string)
	for i := 1; b.locateName(getFooterName(i)) > 0; i++ {
		footerName := getFooterName(i)
		footer := b.getFromName(getFooterName(i))
		footers[i] = footer
		if footer != "" {
			relations[footerName] = b.readPartWithRels(footerName)
		}
	}

	return footers
}

func (b *ZipBuffer) getTempDocumentHeaders(relations map[string]string) map[int]string {
	headers := make(map[int]string)
	for i := 1; b.locateName(getHeaderName(i)) > 0; i++ {
		headerName := getHeaderName(i)
		header := b.getFromName(headerName)
		headers[i] = header
		if header != "" {
			relations[headerName] = b.readPartWithRels(headerName)
		}

	}
	return headers
}

func (b *ZipBuffer) getTempDocumentMainPart(relations map[string]string) (s string) {
	mainPartName := b.getMainPartName()
	relations[mainPartName] = b.readPartWithRels(mainPartName)
	return b.getFromName(mainPartName)
}

func (b *ZipBuffer) getTempDocumentContentTypes(relations map[string]string) string {
	contentTypesName := getDocumentContentTypesName()
	relations[contentTypesName] = b.readPartWithRels(contentTypesName)
	return b.getFromName(contentTypesName)
}

func (b *ZipBuffer) getTempDocumentSettingsPart(relations map[string]string) string {
	settingName := getSettingsPartName()
	relations[settingName] = b.readPartWithRels(settingName)
	return b.getFromName(settingName)

}

func getRelationsName(s string) string {
	return StringBuilder("word/_rels/", filepath.Base(s), ".rels")
}

//定位位置
func (b *ZipBuffer) locateName(s string) int {
	for k, file := range b.files() {
		if file.Name == s {
			return k
		}
	}
	return -1
}

//通过定位读取zip文件
func (b *ZipBuffer) readFileWithIndex(index int) (string, error) {
	if index == -1 {
		return "", nil
	}
	rc, err := b.files()[index].Open()
	if err != nil {
		return "", errors.New("func : readfile , zip文件打开失败")
	}
	content, err := ioutil.ReadAll(rc)
	if err != nil {
		return "", errors.New("func : readfile , zip读取打开失败")
	}
	return StringBuilder(ByteToString(content)), nil
}

//获取内容
func (b *ZipBuffer) getFromName(s string) string {
	index := b.locateName(s)
	res, _ := b.readFileWithIndex(index)
	return res
}

//header名称
func getHeaderName(index int) string {
	return fmt.Sprintf("word/header%d.xml", index)
}

//footer名
func getFooterName(index int) string {
	return fmt.Sprintf("word/footer%d.xml", index)
}

//setting名
func getSettingsPartName() string {
	return "word/settings.xml"
}

//contentTypes 名
func getDocumentContentTypesName() string {
	return "[Content_Types].xml"
}

//主体word 名称
func (b *ZipBuffer) getMainPartName() (mainPartName string) {
	c := b.getFromName(getDocumentContentTypesName())
	reg := regexp.MustCompile(`PartName="\/(word\/document.*?\.xml)" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document\.main\+xml"`)
	res := reg.FindStringSubmatch(c)
	if len(res) > 1 {
		mainPartName = res[1]
	} else {
		mainPartName = "word/document.xml"
	}

	return mainPartName
}

//ByteToString 字节转字符串
func ByteToString(b []byte) string {
	return *(*string)(unsafe.Pointer(&b))
}

//StringBuilder 字符串拼接
func StringBuilder(s ...string) string {
	var buf bytes.Buffer
	for _, v := range s {
		buf.WriteString(v)
	}
	return buf.String()
}
