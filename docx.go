package docx

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"io/ioutil"
	"os"
	"regexp"
	"strings"
)

//ZipData Contains functions to work with data from a zip file
type ZipData interface {
	files() []*zip.File
	close() error
}

//ZipInMemory Type for in memory zip files
type ZipInMemory struct {
	data *zip.Reader
}

func (d ZipInMemory) files() []*zip.File {
	return d.data.File
}

//Since there is nothing to close for in memory, just nil the data and return nil
func (d ZipInMemory) close() error {
	d.data = nil
	return nil
}

//ZipFile Type for zip files read from disk
type ZipFile struct {
	data *zip.ReadCloser
}

func (d ZipFile) files() []*zip.File {
	return d.data.File
}

func (d ZipFile) close() error {
	return d.data.Close()
}

//ReplaceDocx docx 文档
type ReplaceDocx struct {
	zipReader ZipData
	content   string
	links     string
	headers   map[string]string
	footers   map[string]string
}

//Open 打开word
func Open(file string) (*Docx, *ReplaceDocx) {
	r, err := ReadDocxFile(file)
	if err != nil {
		panic(err)
	}
	docx := r.Editable()
	//修复错误标签
	docx.fixBrokenMacros()

	return docx, r
}

//Editable 初始化
func (r *ReplaceDocx) Editable() *Docx {
	return &Docx{
		files:   r.zipReader.files(),
		content: r.content,
		links:   r.links,
		headers: r.headers,
		footers: r.footers,
	}
}

//Close 关闭资源文件
func (r *ReplaceDocx) Close() error {
	return r.zipReader.close()
}

//Docx 结构体
type Docx struct {
	files   []*zip.File
	content string
	links   string
	headers map[string]string
	footers map[string]string
}

//GetContent 获取内容
func (d *Docx) GetContent() string {
	return d.content
}

//SetContent 设置内容
func (d *Docx) SetContent(content string) {
	d.content = content
}

//ReplaceRaw 直接替换
func (d *Docx) ReplaceRaw(oldString string, newString string, num int) {
	d.content = strings.Replace(d.content, oldString, newString, num)
}

//Replace 替换内容
func (d *Docx) Replace(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.content = strings.Replace(d.content, oldString, newString, num)

	return nil
}

//ReplaceLink 替换连接
func (d *Docx) ReplaceLink(oldString string, newString string, num int) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}
	d.links = strings.Replace(d.links, oldString, newString, num)

	return nil
}

//ReplaceHeader 替换head
func (d *Docx) ReplaceHeader(oldString string, newString string) (err error) {
	return replaceHeaderFooter(d.headers, oldString, newString)
}

//ReplaceFooter 替换word foot
func (d *Docx) ReplaceFooter(oldString string, newString string) (err error) {
	return replaceHeaderFooter(d.footers, oldString, newString)
}

//WriteToFile 写入文件
func (d *Docx) WriteToFile(path string) (err error) {
	var target *os.File
	target, err = os.Create(path)
	if err != nil {
		return
	}
	defer target.Close()
	err = d.Write(target)
	return
}

func (d *Docx) Write(ioWriter io.Writer) (err error) {
	w := zip.NewWriter(ioWriter)
	for _, file := range d.files {
		var writer io.Writer
		var readCloser io.ReadCloser

		writer, err = w.Create(file.Name)
		if err != nil {
			return err
		}
		readCloser, err = file.Open()
		if err != nil {
			return err
		}
		if file.Name == "word/document.xml" {
			writer.Write([]byte(d.content))
		} else if file.Name == "word/_rels/document.xml.rels" {
			writer.Write([]byte(d.links))
		} else if strings.Contains(file.Name, "header") && d.headers[file.Name] != "" {
			writer.Write([]byte(d.headers[file.Name]))
		} else if strings.Contains(file.Name, "footer") && d.footers[file.Name] != "" {
			writer.Write([]byte(d.footers[file.Name]))
		} else {
			writer.Write(streamToByte(readCloser))
		}
	}
	w.Close()
	return
}

func replaceHeaderFooter(headerFooter map[string]string, oldString string, newString string) (err error) {
	oldString, err = encode(oldString)
	if err != nil {
		return err
	}
	newString, err = encode(newString)
	if err != nil {
		return err
	}

	for k := range headerFooter {
		headerFooter[k] = strings.Replace(headerFooter[k], oldString, newString, -1)
	}

	return nil
}

//ReadDocxFromMemory 读取zip
func ReadDocxFromMemory(data io.ReaderAt, size int64) (*ReplaceDocx, error) {
	reader, err := zip.NewReader(data, size)
	if err != nil {
		return nil, err
	}
	zipData := ZipInMemory{data: reader}
	return readDocx(zipData)
}

//ReadDocxFile 读取压缩docx文件
func ReadDocxFile(path string) (*ReplaceDocx, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	zipData := ZipFile{data: reader}
	return readDocx(zipData)
}
func readDocx(reader ZipData) (*ReplaceDocx, error) {
	content, err := readText(reader.files())
	if err != nil {
		return nil, err
	}

	links, err := readLinks(reader.files())
	if err != nil {
		return nil, err
	}

	headers, footers, _ := readHeaderFooter(reader.files())
	return &ReplaceDocx{zipReader: reader, content: content, links: links, headers: headers, footers: footers}, nil
}

func readHeaderFooter(files []*zip.File) (headerText map[string]string, footerText map[string]string, err error) {

	h, f, err := retrieveHeaderFooterDoc(files)

	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	headerText, err = buildHeaderFooter(h)
	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	footerText, err = buildHeaderFooter(f)
	if err != nil {
		return map[string]string{}, map[string]string{}, err
	}

	return headerText, footerText, err
}

func buildHeaderFooter(headerFooter []*zip.File) (map[string]string, error) {

	headerFooterText := make(map[string]string)
	for _, element := range headerFooter {
		documentReader, err := element.Open()
		if err != nil {
			return map[string]string{}, err
		}

		text, err := wordDocToString(documentReader)
		if err != nil {
			return map[string]string{}, err
		}

		headerFooterText[element.Name] = text
	}

	return headerFooterText, nil
}

func readText(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveWordDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func readLinks(files []*zip.File) (text string, err error) {
	var documentFile *zip.File
	documentFile, err = retrieveLinkDoc(files)
	if err != nil {
		return text, err
	}
	var documentReader io.ReadCloser
	documentReader, err = documentFile.Open()
	if err != nil {
		return text, err
	}

	text, err = wordDocToString(documentReader)
	return
}

func wordDocToString(reader io.Reader) (string, error) {
	b, err := ioutil.ReadAll(reader)
	if err != nil {
		return "", err
	}
	return string(b), nil
}

func retrieveWordDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/document.xml" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml file not found")
	}
	return
}

func retrieveLinkDoc(files []*zip.File) (file *zip.File, err error) {
	for _, f := range files {
		if f.Name == "word/_rels/document.xml.rels" {
			file = f
		}
	}
	if file == nil {
		err = errors.New("document.xml.rels file not found")
	}
	return
}

func retrieveHeaderFooterDoc(files []*zip.File) (headers []*zip.File, footers []*zip.File, err error) {
	for _, f := range files {

		if strings.Contains(f.Name, "header") {
			headers = append(headers, f)
		}
		if strings.Contains(f.Name, "footer") {
			footers = append(footers, f)
		}
	}
	if len(headers) == 0 && len(footers) == 0 {
		err = errors.New("headers[1-3].xml file not found and footers[1-3].xml file not found")
	}
	return
}

func streamToByte(stream io.Reader) []byte {
	buf := new(bytes.Buffer)
	buf.ReadFrom(stream)
	return buf.Bytes()
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

//SetValue 批量替换
func (d *Docx) SetValue(replaceMap map[string]string) {
	for k, v := range replaceMap {
		d.Replace(k, v, -1)
	}
}

//找到所有标签 并且去皮
func (d *Docx) getVariablesForPart(search string) []string {
	total := []string{}

	reg := regexp.MustCompile(`\$\{(.*?)\}`)
	contentlabel := reg.FindAllStringSubmatch(search, -1)
	for _, v := range contentlabel {
		total = append(total, v[1])
	}
	return total
}

//修复错误标签
func (d *Docx) fixBrokenMacros() {
	re := regexp.MustCompile(`\$(?:\{|[^{$]*\>\{)[^\}\$]*\}`)

	f := func(s string) (src string) {
		re, _ = regexp.Compile(`<[\S\s]+?>`)
		src = re.ReplaceAllString(s, "")
		return
	}
	//修复 content
	d.content = re.ReplaceAllStringFunc(d.content, f)
	//修复 footers
	if len(d.footers) != 0 {
		for k, v := range d.footers {
			d.footers[k] = re.ReplaceAllStringFunc(v, f)
		}
	}
	//修复 headers
	if len(d.headers) != 0 {
		for k, v := range d.headers {
			d.headers[k] = re.ReplaceAllStringFunc(v, f)
		}
	}
}
