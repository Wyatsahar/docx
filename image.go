package docx

import (
	"bytes"
	"errors"
	"fmt"
	"image"
	_ "image/gif"
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

//GIFTYPE
const (
	GIFTYPE = "image/gif"
	BMPTYPE = "image/x-ms-bmp"
	JPGTYPE = "image/jpeg"
	PNGTYPE = "image/png"
)

//ImgValue 结构体
type ImgValue struct {
	Path, Type string
	Width      int
	Height     int
	Search     string
	Replace    string
}

//SetWidth 设置图片宽度
func (i ImgValue) SetWidth(width int) ImgValue {
	i.Width = width
	return i
}

//SetHeight 设置图片高度度
func (i ImgValue) SetHeight(height int) ImgValue {
	i.Height = height
	return i
}

//GetArrangeImage return imgValue
func GetArrangeImage(path string) ImgValue {
	file, err := os.Open(path)
	if err != nil {
		return ImgValue{}
	}
	defer file.Close()
	imgConfig, t, err := image.DecodeConfig(file)

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
func (d *Docx) SetImagesValues(replace string, img ImgValue) {

	// imgTpl = `<w:pict><v:shape type="#_x0000_t75" style="width:{WIDTH};height:{HEIGHT}"><v:imagedata r:id="{RID}" o:title=""/></v:shape></w:pict>`

	//找到所有标签并且去皮
	contentTags := d.getVariablesForPart(d.MainPart)

	// fmt.Println(d.Relations[d.MainPartName])

	for _, tagV := range imgVariablesFilter(contentTags, replace) {
		//整理每个 标签所用到的 height width
		rid := "rId" + strconv.Itoa((strings.Count(d.Relations[d.MainPartName], "<Relationship")))
		img.Search = tagV

		d.addImageToRelations(d.MainPartName, rid, &img)

		// fmt.Println(rid)
	}

	fmt.Println(d.NewImages)

	if len(d.Footers) != 0 {
		for _, footer := range d.Footers {
			//查找所有尾部标签 去皮
			fmt.Println(d.getVariablesForPart(footer))
		}
	}

	if len(d.Headers) != 0 {
		for _, header := range d.Headers {
			//查找所有头部标签 去皮
			d.getVariablesForPart(header)

		}
	}

}

func (d *Docx) addImageToRelations(partFileName, rid string, img *ImgValue) {
	typeTpl := "<Override PartName=\"/word/media/{IMG}\" ContentType=\"image/{EXT}\"/>"
	// relationTpl := "<Relationship Id=\"{RID}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/{IMG}\"/>"
	// newRelationsTpl := "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>"
	// newRelationsTypeTpl := "<Override PartName=\"/{RELS}\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"

	img.Replace = `word/media/` + `image_` + rid + `_` + pathInfo(partFileName) + `.` + img.Type
	d.NewImages[img.Search] = *img

	typeTpl = strReplace([]string{`{IMG}`, `{EXT}`}, []string{img.Replace, img.Type}, typeTpl)

	d.ContentTypes = strings.Replace(d.ContentTypes, `</Types>`, typeTpl, -1) + `</Types>`

	fmt.Println()
	// fmt.Println(d.ContentTypes)

}

func pathInfo(fileFullName string) string {
	filenameall := path.Base(fileFullName)
	filesuffix := path.Ext(fileFullName)
	return filenameall[0 : len(filenameall)-len(filesuffix)]
}

//获取 标签上设置的图片参数
func getImageArgs(varNameWithArgs string) (varInlineArgs map[string]string) {
	varInlineArgs = make(map[string]string)
	vn := strings.Split(varNameWithArgs, ":")[1:]
	reg := regexp.MustCompile(`([0-9]*[a-z%]{0,2}|auto)x([0-9]*[a-z%]{0,2}|auto)`)
	for k, v := range vn {
		if strings.Index(v, "=") != -1 { // arg=value
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
				break
			case 1:
				varInlineArgs["height"] = v
				break
			case 2:
				varInlineArgs["ratio"] = v
				break
			}
		}
	}

	return
}

func listString(s, spe string) (argKey, argValue string) {
	arg := strings.SplitN(s, spe, 2)
	return arg[0], arg[1]
}

//搜索符合标准的图片标签
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

//获取图片信息
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
	d.MainPart = re.ReplaceAllStringFunc(d.MainPart, f)
	//修复 footers
	if len(d.Footers) != 0 {
		for k, footer := range d.Footers {
			d.Footers[k] = re.ReplaceAllStringFunc(footer, f)
		}
	}
	//修复 headers
	if len(d.Headers) != 0 {
		for k, header := range d.Headers {
			d.Headers[k] = re.ReplaceAllStringFunc(header, f)
		}
	}
}
