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
)

// Config 配置项
type Config struct {
	PlaceholderPrefix string // 占位符前缀，如 {{
	PlaceholderSuffix string // 占位符后缀，如 }}
}

var DefaultConfig = Config{
	PlaceholderPrefix: "{{",
	PlaceholderSuffix: "}}",
}

// Docx 文档
type Docx struct {
	Path             string
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
	Config           Config
}

// ZipData Contains functions to work with data from a zip file
type ZipData interface {
	files() []*zip.File
	close() error
}

// ZipBuffer zip buffer
type ZipBuffer struct {
	reader *zip.Reader
	closer io.Closer
}

func (b *ZipBuffer) files() []*zip.File {
	return b.reader.File
}

// Close 关闭底层的 zip reader
func (b *ZipBuffer) Close() error {
	if b.closer != nil {
		return b.closer.Close()
	}
	return nil
}

// Close 关闭资源，建议在 Docx 使用完毕后手动调用
func (d *Docx) Close() error {
	if d.ZipBuffer != nil {
		return d.ZipBuffer.Close()
	}
	return nil
}

// Load 初始化Docx (兼容旧接口，默认使用 DefaultConfig)
func Load(path string) (*Docx, error) {
	return LoadWithOptions(path, DefaultConfig)
}

// LoadWithOptions 带配置初始化Docx
func LoadWithOptions(path string, config Config) (*Docx, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}
	fi, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, err
	}
	d, err := LoadFromReader(f, fi.Size(), config)
	if err != nil {
		f.Close()
		return nil, err
	}
	d.Path = path
	d.ZipBuffer.closer = f // 确保文件的关闭
	return d, nil
}

// LoadFromReader 从 Reader 加载文档
func LoadFromReader(r io.ReaderAt, size int64, config Config) (*Docx, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("failed to create zip reader: %w", err)
	}

	b := &ZipBuffer{reader: zr}

	Relations := make(map[string]string)
	Headers := b.getTempDocumentHeaders(Relations)
	Footers := b.getTempDocumentFooters(Relations)
	MainPartName, MainPart := b.getTempDocumentMainPart(Relations)
	SettingsPartName, SettingsPart := b.getTempDocumentSettingsPart(Relations)
	ContentTypesName, ContentTypes := b.getTempDocumentContentTypes(Relations)

	d := &Docx{
		ZipBuffer:        b,
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
		Config:           config,
	}

	d.fixBrokenMacros()
	return d, nil
}

// Save 保存文件
func (d *Docx) save() error {
	return d.SaveToFile(d.Path)
}

// SaveToFile 另存为
func (d *Docx) SaveToFile(path string) (err error) {
	w, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("failed to create file: %w", err)
	}
	defer w.Close()
	_, err = d.WriteTo(w)
	return err
}

// WriteTo 将文档写入指定 Writer，符合 io.WriterTo 接口
func (d *Docx) WriteTo(w io.Writer) (int64, error) {
	cw := &countingWriter{w: w}
	wr := zip.NewWriter(cw)
	defer wr.Close()

	for _, file := range d.ZipBuffer.files() {
		xmlString := d.ZipBuffer.getFromName(file.Name)
		for headerIndex, header := range d.Headers {
			if file.Name == getHeaderName(headerIndex) {
				xmlString = header
			}
		}

		for footerIndex, footer := range d.Footers {
			if file.Name == getFooterName(footerIndex) {
				xmlString = footer
			}
		}

		if file.Name == d.SettingsPartName {
			xmlString = d.SettingsPart
		}

		if file.Name == d.MainPartName && d.MainPart != "" {
			xmlString = d.MainPart
		}

		if file.Name == d.ContentTypesName && d.ContentTypes != "" {
			xmlString = d.ContentTypes
		}

		err := d.savePartWithRels(wr, file.Name, xmlString)
		if err != nil {
			return cw.count, fmt.Errorf("failed to save part %s: %w", file.Name, err)
		}
	}

	// 写入图片文件
	if len(d.NewImages) > 0 {
		err := d.saveImages(wr)
		if err != nil {
			return cw.count, fmt.Errorf("failed to save images: %w", err)
		}
	}

	wr.Close()
	return cw.count, nil
}

type countingWriter struct {
	w     io.Writer
	count int64
}

func (cw *countingWriter) Write(p []byte) (n int, err error) {
	n, err = cw.w.Write(p)
	cw.count += int64(n)
	return
}

// SaveToBuffer 保存为内存 Buffer
func (d *Docx) SaveToBuffer() (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	_, err := d.WriteTo(buf)
	return buf, err
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
		f, err := os.Open(image.Path)
		if err != nil {
			return fmt.Errorf("failed to open image %s: %w", image.Path, err)
		}
		imageContent, err := ioutil.ReadAll(f)
		f.Close() // 显式关闭图片文件
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
		return nil
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
			return errors.New("create rels error")
		}
		_, err = relsWriter.Write([]byte(v))
		if err != nil {
			return errors.New("write rels error")
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
	if reflect.TypeOf(s[0]).Kind() == reflect.Map {
		m, ok := s[0].(map[string]string)
		if !ok {
			return errors.New("map参数类型错误，应为 map[string]string")
		}
		for search, replace := range m {
			d.replace(search, replace, -1)
		}
	} else if len(s) == 2 && reflect.TypeOf(s[0]).Kind() == reflect.String && reflect.TypeOf(s[1]).Kind() == reflect.String {
		d.replace(s[0].(string), s[1].(string), -1)
	} else {
		return errors.New("参数类型错误")
	}

	return nil
}

// replace 替换文本
func (d *Docx) replace(search, replace string, limit int) error {
	encodeSearch, err := encode(StringBuilder(d.Config.PlaceholderPrefix, search, d.Config.PlaceholderSuffix))
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
	for i := 1; b.locateName(getFooterName(i)) >= 0; i++ {
		footerName := getFooterName(i)
		footer := b.getFromName(footerName)
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
	for i := 1; b.locateName(getHeaderName(i)) >= 0; i++ {
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
		relations[contentTypesName] = typeRelsContent
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

// 定位位置
func (b *ZipBuffer) locateName(s string) int {
	for k, file := range b.files() {
		if file.Name == s {
			return k
		}
	}
	return -1
}

// 通过定位读取zip文件
func (b *ZipBuffer) readFileWithIndex(index int) (string, error) {
	if index == -1 {
		return "", nil
	}
	rc, err := b.files()[index].Open()
	if err != nil {
		return "", fmt.Errorf("failed to open zip file member: %w", err)
	}
	defer rc.Close()
	content, err := ioutil.ReadAll(rc)
	if err != nil {
		return "", fmt.Errorf("failed to read zip file member: %w", err)
	}
	return string(content), nil
}

// 获取内容
func (b *ZipBuffer) getFromName(s string) string {
	index := b.locateName(s)
	res, _ := b.readFileWithIndex(index)
	return res
}

// header名称
func getHeaderName(index int) string {
	return fmt.Sprintf("word/header%d.xml", index)
}

// footer名
func getFooterName(index int) string {
	return fmt.Sprintf("word/footer%d.xml", index)
}

// setting名
func getSettingsPartName() string {
	return "word/settings.xml"
}

// contentTypes 名
func getDocumentContentTypesName() string {
	return "[Content_Types].xml"
}

// 主体word 名称
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

// ByteToString 字节转字符串 (保留函数名以维持兼容性，但移除 unsafe)
func ByteToString(b []byte) string {
	return string(b)
}

// StringBuilder 字符串拼接，优化性能
func StringBuilder(s ...string) string {
	var sb strings.Builder
	for _, v := range s {
		sb.WriteString(v)
	}
	return sb.String()
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
