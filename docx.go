package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
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
	ZipBuffer        *ZipBuffer
	MainPart         string
	MainPartName     string
	SettingsPart     string
	SettingsPartName string
	ContentTypes     string
	ContentTypesName string
	Headers          map[int]string
	Footers          map[int]string
	Relations        map[string]string
	NewImages        map[string]ImgValue
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
	d, rc := getDocx(path)
	//整理错误标签
	d.fixBrokenMacros()
	return d, rc
}

func getDocx(path string) (*Docx, *ZipBuffer) {
	//打开zip文件
	rc, _ := zip.OpenReader(path)

	//把zip复制一份
	b := ZipBuffer{rc}

	Relations := make(map[string]string)
	Headers := b.getTempDocumentHeaders(Relations)
	Footers := b.getTempDocumentFooters(Relations)
	MainPartName, MainPart := b.getTempDocumentMainPart(Relations)
	SettingsPartName, SettingsPart := b.getTempDocumentSettingsPart(Relations)
	ContentTypesName, ContentTypes := b.getTempDocumentContentTypes(Relations)

	return &Docx{
		ZipBuffer:        &b,
		Headers:          Headers,
		Footers:          Footers,
		Relations:        Relations,
		MainPartName:     MainPartName,
		MainPart:         MainPart,
		SettingsPart:     SettingsPart,
		SettingsPartName: SettingsPartName,
		ContentTypes:     ContentTypes,
		ContentTypesName: ContentTypesName,
		NewImages:        make(map[string]ImgValue),
	}, &b
}

//SaveToFile 保存文件
func (d *Docx) SaveToFile(path string) (err error) {

	w, err := os.Create(path)
	wr := zip.NewWriter(w)
	if err != nil {
		return err
	}
	for _, file := range d.ZipBuffer.files() {
		xml := d.ZipBuffer.getFromName(file.Name)
		for headerIndex, header := range d.Headers {
			if file.Name == getHeaderName(headerIndex) {
				xml = header
			}
		}

		for footerIndex, footer := range d.Footers {
			if file.Name == getFooterName(footerIndex) {
				xml = footer
			}
		}

		if file.Name == d.SettingsPartName {
			xml = d.SettingsPart
		}

		if file.Name == d.MainPartName && d.MainPart != "" {
			xml = d.MainPart
		}

		if file.Name == d.ContentTypes && d.ContentTypes != "" {
			xml = d.ContentTypes
		}

		err := d.savePartWithRels(wr, file.Name, xml)
		if err != nil {
			fmt.Println(err.Error())
		}
	}

	//写入图片文件
	if len(d.NewImages) > 0 {
		_ = d.saveImages(wr)
	}

	_ = wr.Close()
	w.Close()
	return nil
}

func (d *Docx) saveImages(wr *zip.Writer) error {
	var repetition = make([]string, 0, len(d.NewImages))

	for _, image := range d.NewImages {
		//在切片中查找重复
		if findStrInSlice(repetition, image.Replace) != -1 {
			continue
		}

		writer, err := wr.Create(`word/media/` + image.Replace)
		if err != nil {
			return err
		}
		files, err := os.Open(image.Path)
		if files == nil{
			return err
		}
		defer files.Close()
		if err != nil {
			return err
		}
		imageContent, err := ioutil.ReadAll(files)
		if err != nil {
			return err
		}
		_, err = writer.Write(imageContent)
		if err != nil {
			return err
		}
		repetition = append(repetition, image.Replace)
	}

	return nil
}

func (d *Docx) savePartWithRels(wr *zip.Writer, filename, xml string) (err error) {

	if _, ok := d.Relations[getRemoveRelationsName(filename)]; ok && getRemoveRelationsName(filename) != filename {
		return
	}

	writer, err := wr.Create(filename)
	if err != nil {
		return err
	}

	_, err = writer.Write([]byte(xml))
	if err != nil {
		return err
	}

	if v, ok := d.Relations[filename]; ok {
		relsFileName := getRelationsName(filename)

		relsWriter, err := wr.Create(relsFileName)
		if err != nil {
			return errors.New("create error")
		}
		_, err = relsWriter.Write([]byte(v))
		if err != nil {
			return errors.New("relsWriter error")
		}
	}
	return nil
}

/*
SetValue 替换文本
	(d *Docx) SetValue( map[search]replace )
	(d *Docx) SetValue( search string, replace string)
*/
func (d *Docx) SetValue(s ...interface{}) error {
	if len(s) != 2 && len(s) != 1 {
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
	encodeSearch, err := encode(StringBuilder("${", search, "}"))
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
		if footer == "" {
			delete(relations, footerName)
		} else {
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
		if header == "" {
			delete(relations, headerName)
		} else {
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
	} else {
		relations[contentTypesName] = b.readPartWithRels(contentTypesName)
	}

	return contentTypesName, b.getFromName(contentTypesName)
}

func (b *ZipBuffer) getTempDocumentSettingsPart(relations map[string]string) (name, s string) {
	settingName := getSettingsPartName()
	settingRelsContent := b.readPartWithRels(settingName)
	if settingRelsContent == "" {
		delete(relations, settingName)
	} else {
		relations[settingName] = settingRelsContent
	}

	return settingName, b.getFromName(settingName)

}

func getRelationsName(s string) string {
	return StringBuilder("word/_rels/", filepath.Base(s), ".rels")
}

// word/_rels/aaa.xml.rels
func getRemoveRelationsName(s string) string {
	return strReplace([]string{"_rels/", ".rels"}, []string{"", ""}, s)
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

func findStrInSlice(slice []string, val string) int {
	for i, item := range slice {
		if item == val {
			return i
		}
	}
	return -1
}

func strReplace(search, replace []string, s string) string {
	if len(search) != len(replace) {
		return ""
	}

	for i := 0; i < len(search); i++ {
		s = strings.ReplaceAll(s, search[i], replace[i])
	}

	return s
}
