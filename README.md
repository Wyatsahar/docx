

## A simple go (golang) Microsoft Word (. Docx) tool library to replace images/text

### 一个简单的golang Word操作库
### 参照 PhpOffice/PhpWord 写的一个小工具

#### 替换文本

```go


import (
	"github.com/wyatsahar/docx"
)

func main() {
	//载入word
	doc := docx.Load("./document_test.docx")
	//批量文本替换
	var data = make(map[string]string)
	data["search"] = "批量替换"
	data["search1"] = "批量替换1"
	doc.SetValue(data)
	//单独文本替换
	doc.SetValue("search2", "单独替换")
        //另存为
	doc.SaveToFile("./new_result_2.docx")
	
}

```

#### 复制行

​	CloneRow(mark,num)


| 编号 | 姓名 |
| - | - |
| ${id} | ${name} |

```go
import (
	"github.com/wyatsahar/docx"
)

func main() {
        //载入word
        doc := docx.Load("./document_test.docx")
        //复制行
        doc.CloneRow("id", 3)
        //替换复制行后的标签
        var data1 = make(map[string]string)
        data1["id#0"] = "1"
        data1["name#0"] = "张三"
        data1["id#1"] = "2"
        data1["name#1"] = "李四"
        data1["id#2"] = "3"
        data1["name#2"] = "王五"
        doc.SetValue(data1)
        //另存为
        doc.SaveToFile("./new_result_2.docx")

	
}
```


| 编号 | 姓名 |
| - | - |
| 1 | 张三 |
| 2 | 李四 |
| 3 | 王五 |



#### 替换图片

```go
import (
	"github.com/wyatsahar/docx"
)

func main (){
    //载入word
	doc := docx.Load("./document_test.docx")
    
    //增加图片 设置宽高
	img1 := doc.GetArrangeImage("./aaa.png").SetHeight(100).SetWidth(100)
	img2 := doc.GetArrangeImage("./bbb.jpg").SetHeight(150).SetWidth(200)
	//替换图片
	doc.SetImagesValues("img", img1)
	doc.SetImagesValues("img3", img2)
	doc.SetImagesValues("img2", img1)
        //另存为
        doc.SaveToFile("./new_result_2.docx")

}
```

