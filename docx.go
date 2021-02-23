package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"unsafe"
)

//Docx 文档
type Docx struct {
	files            []*zip.File
	MainPart         string
	MainPartName     string
	SettingsPart     string
	SettingsPartName string
	ContentTypes     string
	ContentTypesName string
	Headers          map[int]string
	Footers          map[int]string
	Relations        map[string]string
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

func (b *ZipBuffer) files() []*zip.File {
	return b.rc.File
}

//Close 关闭
func (b *ZipBuffer) Close() error {
	return b.rc.Close()
}

//LoadInit 初始化Docx
func LoadInit(path string) (*Docx, *ZipBuffer) {
	// func LoadInit(path string) {
	//打开zip文件
	rc, _ := zip.OpenReader(path)

	b := ZipBuffer{rc}
	Relations := make(map[string]string)
	Headers := b.getTempDocumentHeaders(Relations)
	Footers := b.getTempDocumentFooters(Relations)
	MainPartName, MainPart := b.getTempDocumentMainPart(Relations)
	SettingsPartName, SettingsPart := b.getTempDocumentSettingsPart(Relations)
	ContentTypesName, ContentTypes := b.getTempDocumentContentTypes(Relations)

	return &Docx{
		files:            b.files(),
		Headers:          Headers,
		Footers:          Footers,
		Relations:        Relations,
		MainPartName:     MainPartName,
		MainPart:         MainPart,
		SettingsPart:     SettingsPart,
		SettingsPartName: SettingsPartName,
		ContentTypes:     ContentTypes,
		ContentTypesName: ContentTypesName,
	}, &b
}

//SaveToFile 保存文件
func (d *Docx) SaveToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	wr := zip.NewWriter(target)
	if err != nil {
		return err
	}

	for k, v := range d.files {
		fmt.Println(k)
		fmt.Println(v)
	}

	for index, header := range d.Headers {
		d.savePartWithRels(wr, getHeaderName(index), header)
	}

	d.savePartWithRels(wr, d.MainPartName, d.MainPart)
	d.savePartWithRels(wr, d.SettingsPartName, d.SettingsPart)
	d.savePartWithRels(wr, d.ContentTypesName, d.ContentTypes)

	for index, footer := range d.Footers {
		d.savePartWithRels(wr, getFooterName(index), footer)
	}

	wr.Close()
	return nil
}

func (d *Docx) savePartWithRels(wr *zip.Writer, filename, xml string) (err error) {
	writer, err := wr.Create(filename)
	if err != nil {
		return err
	}
	io.WriteString(writer, xml)
	if _, ok := d.Relations[filename]; ok {
		relsFileName := getRelationsName(filename)
		relsWriter, err := wr.Create(relsFileName)
		if err != nil {
			return err
		}
		io.WriteString(relsWriter, d.Relations[filename])
	}
	return nil
}

/*
SetValue 替换文本
	(d *Docx) SetValue( map[search]replace )
	(d *Docx) SetValue( search string, replace string)
*/
func (d *Docx) SetValue(s ...interface{}) error {
	if len(s) != 2 {
		return errors.New("参数长度错误")
	}
	//如果第一个参数为map
	// 使用  map[string]string 来替换
	if reflect.TypeOf(s[0]).Kind() == reflect.Map {
		for search, replace := range s[0].(map[string]string) {
			d.replace(search, replace, -1)
		}
	} else if reflect.TypeOf(s[0]).Kind() == reflect.String && reflect.TypeOf(s[1]).Kind() == reflect.String {
		d.replace(s[0].(string), s[1].(string), -1)
	} else {
		return errors.New("参数错误")
	}

	return nil
}

//replace 替换文本
func (d *Docx) replace(search, replace string, limit int) error {
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

	d.MainPart = strings.Replace(d.MainPart, search, replace, limit)

	for headerIndex, header := range d.Headers {
		d.Headers[headerIndex] = strings.Replace(header, search, replace, limit)
	}
	for footerIndex, footer := range d.Footers {
		d.Footers[footerIndex] = strings.Replace(footer, search, replace, limit)
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
			delete(relations, headerName)
			relations[headerName] = b.readPartWithRels(headerName)
		}

	}
	return headers
}

func (b *ZipBuffer) getTempDocumentMainPart(relations map[string]string) (name, s string) {
	mainPartName := b.getMainPartName()
	relations[mainPartName] = b.readPartWithRels(mainPartName)
	return mainPartName, b.getFromName(mainPartName)
}

func (b *ZipBuffer) getTempDocumentContentTypes(relations map[string]string) (name, s string) {
	contentTypesName := getDocumentContentTypesName()
	typeRelsContent := b.readPartWithRels(contentTypesName)
	if typeRelsContent == "" {
		delete(relations, contentTypesName)
		relations[contentTypesName] = b.readPartWithRels(contentTypesName)
	}

	return contentTypesName, b.getFromName(contentTypesName)
}

func (b *ZipBuffer) getTempDocumentSettingsPart(relations map[string]string) (name, s string) {
	settingName := getSettingsPartName()
	settingRelsContent := b.readPartWithRels(settingName)
	if settingRelsContent == "" {
		delete(relations, settingName)
		relations[settingName] = settingRelsContent
	}

	return settingName, b.getFromName(settingName)

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
