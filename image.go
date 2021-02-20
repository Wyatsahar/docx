package docx

import (
	"bytes"
	"errors"
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
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
	Path   string
	Weight uint
	Height uint
}

// SetImagesValues 设置图片
func (d *Docx) SetImagesValues(replace map[string]string) {

	// imgTpl = `<w:pict><v:shape type="#_x0000_t75" style="width:{WIDTH};height:{HEIGHT}"><v:imagedata r:id="{RID}" o:title=""/></v:shape></w:pict>`

	//找到所有标签并且去皮
	contentTags := d.getVariablesForPart(d.content)
	for k, v := range replace {
		for _, tagV := range imgVariablesFilter(contentTags, k) {
			//整理每个 标签所用到的 height width
			// fmt.Println(v)    //图片路径
			// fmt.Println(tagV) //标签所匹配的属性
			imageType, err := getImagesType(v)
			if err != nil {
				d.Replace(`${`+tagV+`}`, err.Error(), -1)
				break
			}

			im := handleImagesArg(v, imageType, getImageArgs(tagV))

			fmt.Println(k)
			fmt.Println(im)
		}
	}
	if len(d.footers) != 0 {
		for _, footer := range d.footers {
			//查找所有尾部标签 去皮
			d.getVariablesForPart(footer)
		}
	}

	if len(d.headers) != 0 {
		for _, header := range d.headers {
			//查找所有头部标签 去皮
			d.getVariablesForPart(header)

		}
	}

}

//整理图片数据
func handleImagesArg(imgPath, imgType string, imgArt map[string]string) map[string]string {
	imgArt["src"] = imgPath
	imgArt["type"] = imgType
	imgArt["height"] = imgArt["height"] + "px"
	imgArt["width"] = imgArt["width"] + "px"
	return imgArt
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
