package docx

import (
	"bytes"
	"errors"
	"image"
	_ "image/gif" // 检测图片类型
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"os"
	"path"
	"regexp"
	"strconv"
	"strings"
)

var gifByte = []byte("GIF")
var bmpByte = []byte("BM")
var jpgByte = []byte{0xff, 0xd8, 0xff}
var pngByte = []byte{0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a}

// GIFTYPE
const (
	GIFTYPE = "image/gif"
	BMPTYPE = "image/x-ms-bmp"
	JPGTYPE = "image/jpeg"
	PNGTYPE = "image/png"
)

// ImgValue 结构体
type ImgValue struct {
	Path, Type string
	Width      int
	Height     int
	Search     string
	Replace    string
	Rid        string
}

// SetWidth 设置图片宽度
func (i ImgValue) SetWidth(width int) ImgValue {
	i.Width = width
	return i
}

// SetHeight 设置图片高度度
func (i ImgValue) SetHeight(height int) ImgValue {
	i.Height = height
	return i
}

// GetArrangeImage return imgValue
func (d *Docx) GetArrangeImage(path string) ImgValue {
	file, err := os.Open(path)
	if err != nil {
		return ImgValue{}
	}
	defer file.Close()
	imgConfig, t, err := image.DecodeConfig(file)
	if err != nil {
		return ImgValue{}
	}
	width := imgConfig.Width
	height := imgConfig.Height

	return ImgValue{
		Path:   path,
		Width:  width,
		Height: height,
		Type:   t,
	}
}

// SetImagesValues 设置图片
/*
	图片 header document foot 写的位置都是不一样的
	所以只有在设置的时候能查到
*/
func (d *Docx) SetImagesValues(search string, img ImgValue) {

	if img.Type == "" {
		return
	}

	//找到所有标签并且去皮
	contentTags := d.getVariablesForPart(d.MainPart)
	d.MainPart = d.addImageToDocx(contentTags, search, img, d.MainPartName, d.MainPart)

	for headerName, header := range d.Headers {
		//找到所有标签并且去皮
		contentTags1 := d.getVariablesForPart(header)
		if len(contentTags1) > 0 {
			hs := d.addImageToDocx(contentTags1, search, img, "word/header"+strconv.Itoa(headerName)+".xml", header)
			if hs != "" {
				d.Headers[headerName] = hs
			}

		}

	}

	for footerName, footer := range d.Footers {
		//找到所有标签并且去皮
		contentTags2 := d.getVariablesForPart(footer)
		if len(contentTags2) > 0 {
			fs := d.addImageToDocx(contentTags2, search, img, "word/footer"+strconv.Itoa(footerName)+".xml", footer)
			if fs != "" {
				d.Headers[footerName] = fs
			}
		}
	}
}

func (d *Docx) addImageToDocx(contentTags []string, search string, img ImgValue, fileName string, content string) string {
	imgTpl := `<w:pict><v:shape type="#_x0000_t75" style="width:{WIDTH};height:{HEIGHT}"><v:imagedata r:id="{RID}" o:title=""/></v:shape></w:pict>`
	for _, mark := range imgVariablesFilter(contentTags, search) {
		//整理每个 标签所用到的 height width
		img.Search = mark
		rid := d.getRid(fileName, &img)
		d.addImageToRelations(fileName, rid, &img)

		xmlImage := strReplace([]string{`{RID}`, `{WIDTH}`, `{HEIGHT}`}, []string{`rId` + rid, strconv.Itoa(img.Width), strconv.Itoa(img.Height)}, imgTpl)

		re := regexp.MustCompile(`(<[^<]+>)([^<]*)(` + regexp.QuoteMeta(ensureMacroCompleted(d, mark)) + `)([^>]*)(<[^>]+>)`)

		matches := re.FindStringSubmatch(content)

		if len(matches) > 0 {
			wholeTag := matches[0]

			matches = matches[1:]
			openTag := matches[0]
			prefix := matches[1]
			postfix := matches[3]
			closeTag := matches[4]

			replacexml := StringBuilder(openTag, prefix, closeTag, xmlImage, openTag, postfix, closeTag)

			return strings.Replace(content, wholeTag, replacexml, -1)

		}

	}
	return ""
}

func (d *Docx) getRid(partFileName string, img *ImgValue) string {

	rid := strconv.Itoa(strings.Count(d.Relations[partFileName], "<Relationship"))

	return rid
}

/*
	map["img:100:200"]{
		Path:"aaa.png"
		Width:100,
		Height,200,
		Search:img:100:200,
		Replace:document_rid1.png
	}
*/
func (d *Docx) addImageToRelations(partFileName string, rid string, img *ImgValue) {
	typeTpl := "<Override PartName=\"/word/media/{IMG}\" ContentType=\"image/{EXT}\"/>"
	relationTpl := "<Relationship Id=\"{RID}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{IMG}\"/>"
	newRelationsTpl := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
	newRelationsTypeTpl := "<Override PartName=\"/{RELS}\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"

	if _, ok := d.NewImages[img.Search]; !ok && !d.findDuplicateTags(*img) {
		partName := pathInfo(partFileName)
		img.Rid = rid
		img.Replace = `image_` + rid + `_` + partName + `.` + img.Type
		d.NewImages[img.Search] = *img

		typeTpl = strReplace([]string{`{IMG}`, `{EXT}`}, []string{img.Replace, img.Type}, typeTpl)
		d.ContentTypes = strings.Replace(d.ContentTypes, `</Types>`, typeTpl, -1) + `</Types>`
	} else {
		d.NewImages[img.Search] = d.getDuplicateTags(*img)
	}
	xmlImageRelation := strReplace([]string{`{RID}`, `{IMG}`}, []string{"rId" + rid, d.NewImages[img.Search].Replace}, relationTpl)

	//如果没有 则添加
	if _, ok := d.Relations[partFileName]; !ok {
		d.Relations[partFileName] = newRelationsTpl
		xmlRelationsType := strings.Replace(newRelationsTypeTpl, `{RELS}`, getRelationsName(partFileName), -1) + `</Types>`

		d.ContentTypes = strings.Replace(d.ContentTypes, `</Types>`, xmlRelationsType, -1) + `</Types>`
	}

	d.Relations[partFileName] = strings.Replace(d.Relations[partFileName], `</Relationships>`, xmlImageRelation, -1) + `</Relationships>`
}

// 获取一样的
func (d *Docx) getDuplicateTags(img ImgValue) ImgValue {
	for _, v := range d.NewImages {
		if v.Path == img.Path {
			return v
		}
	}
	return ImgValue{}
}

func (d *Docx) findDuplicateTags(img ImgValue) (r bool) {
	r = false
	for _, v := range d.NewImages {
		if v.Path == img.Path {
			r = true
			break
		}
	}

	return
}

func pathInfo(fileFullName string) string {
	filenameall := path.Base(fileFullName)
	filesuffix := path.Ext(fileFullName)
	return filenameall[0 : len(filenameall)-len(filesuffix)]
}

// 获取 标签上设置的图片参数
func getImageArgs(varNameWithArgs string) (varInlineArgs map[string]string) {
	varInlineArgs = make(map[string]string)
	vn := strings.Split(varNameWithArgs, ":")[1:]
	reg := regexp.MustCompile(`([0-9]*[a-z%]{0,2}|auto)x([0-9]*[a-z%]{0,2}|auto)`)
	for k, v := range vn {
		if strings.Contains(v, "=") { // arg=value
			argName, argValue := listString(v, "=")
			if argName == "size" {
				varInlineArgs["width"], varInlineArgs["height"] = listString(argValue, "x")
			} else {
				varInlineArgs[strings.ToLower(argName)] = argValue
			}
		} else if reg.MatchString(v) { // 60x40
			varInlineArgs["width"], varInlineArgs["height"] = listString(v, "x")
		} else {
			switch k {
			case 0:
				varInlineArgs["width"] = v
			case 1:
				varInlineArgs["height"] = v
			case 2:
				varInlineArgs["ratio"] = v
			}
		}
	}

	return
}

func listString(s, spe string) (argKey, argValue string) {
	arg := strings.SplitN(s, spe, 2)
	return arg[0], arg[1]
}

// 搜索符合标准的图片标签
func imgVariablesFilter(variables []string, check string) (variable []string) {
	f := func(variable string, searchString string) bool {

		restr, err := regexp.Compile(`^` + searchString + `:`)
		if err != nil {
			panic("file : docx/image.go func:f(variable string,searchString string) err:regexp.Compile return err")
		}
		re := regexp.MustCompile(restr.String())

		return variable == searchString || re.MatchString(variable)
	}

	for _, v := range variables {
		if f(v, check) {
			variable = append(variable, v)
		}
	}
	return
}

// 获取图片信息
func getImagesType(imgName string) (string, error) {
	fi, err := ioutil.ReadFile(imgName)
	if err != nil {
		errMes := ""
		if os.IsNotExist(err) {
			errMes = imgName + " 不存在"
		} else if os.IsPermission(err) {
			errMes = imgName + " 没有权限"
		} else {
			errMes = imgName + " 读取失败"
		}

		return "", errors.New(errMes)
	}
	var itype string
	if bytes.Equal(pngByte, fi[0:8]) {
		itype = PNGTYPE
	}
	if bytes.Equal(gifByte, fi[0:3]) {
		itype = GIFTYPE
	}
	if bytes.Equal(bmpByte, fi[0:2]) {
		itype = BMPTYPE
	}
	if bytes.Equal(jpgByte, fi[0:3]) {
		itype = JPGTYPE
	}

	if itype == "" {
		return itype, errors.New("undefined type")
	}
	return itype, nil
}

// 找到所有标签 并且去皮
func (d *Docx) getVariablesForPart(search string) []string {
	var total []string

	prefix := regexp.QuoteMeta(d.Config.PlaceholderPrefix)
	suffix := regexp.QuoteMeta(d.Config.PlaceholderSuffix)
	reg := regexp.MustCompile(prefix + `(.*?)` + suffix)
	contentlabel := reg.FindAllStringSubmatch(search, -1)
	for _, v := range contentlabel {
		total = append(total, v[1])
	}
	return total
}

// 修复错误标签 (防止标签被 Word 内部 XML 标签切断)
func (d *Docx) fixBrokenMacros() {
	if len(d.Config.PlaceholderPrefix) < 2 {
		return
	}
	p1 := regexp.QuoteMeta(d.Config.PlaceholderPrefix[:1])
	pRest := regexp.QuoteMeta(d.Config.PlaceholderPrefix[1:])
	s := regexp.QuoteMeta(d.Config.PlaceholderSuffix)

	// 匹配前缀第一位 + (前缀剩余位 或 标签+前缀剩余位) + 内容 + 后缀
	regStr := p1 + `(?:` + pRest + `|[\s\S]*?\>` + pRest + `)[\s\S]*?` + s
	re, err := regexp.Compile(regStr)
	if err != nil {
		return
	}

	f := func(s string) (src string) {
		cleanReg, _ := regexp.Compile(`<[\S\s]+?>`)
		src = cleanReg.ReplaceAllString(s, "")
		return
	}
	// 修复 content
	d.MainPart = re.ReplaceAllStringFunc(d.MainPart, f)
}
