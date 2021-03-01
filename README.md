## A simple go (golang) Microsoft Word (. Docx) tool library

### 一个简单的golang Word操作库

使用例子

```go

import (
	"github.com/wyatsahar/docx"
)

func main() {
	d, rc := docx.LoadInit("./test.docx")
    
    //批量替换 ${search}
	var a = make(map[string]string)
	a["search"] = "replace"
    d.SetValue(a)
    //单独替换  ${search1}
	d.SetValue("search1", "replace1")

	d.SaveToFile("./new_result_2.docx")

	rc.Close()
}

```

增加cloneRow


| 编号 | 姓名 | 年龄 |
| - | - | - |
| ${id} | ${name} | ${age} |

```go
import (
	"github.com/wyatsahar/docx"
)

func main() {
	r, _ := docx.ReadDocxFile("test.docx")
	b := r.Editable()
	b.CloneRow("${id}", 3)
//	b.Replace(`${img#0}`, "test", -1)
//	b.Replace(`${img#1}`, "test1", -1)
//	b.Replace(`${img#2}`, "test2", -1)
	b.WriteToFile("./new_result_2.docx")
	r.Close()
}
```


| 编号 | 姓名 | 年龄 |
| - | - | - |
| ${id#0} | ${name#0} | ${age#0} |
| ${id#1} | ${name#1} | ${age#1} |
| ${id#2} | ${name#2} | ${age#2} |
