## go-excels库，提供excel模板导入和模板导出的简单功能

### 导入功能，读取一个excel文件，返回一个结果集
![image](https://github.com/hq0819/go-excels/assets/52909899/38d29572-d2f8-4567-8b17-fcd8930715e0)
```go

func main() {
	excel := excelTemplate.ReaderExcel[Person](`C:\Users\heqin11\Desktop\aaa.xlsx`, "")
	fmt.Println(excel)
}

type Person struct {
	Name    string `excelTemplate:"姓名"`
	Age     int64  `excelTemplate:"年龄"`
	Address string `excelTemplate:"地址"`
}

```
![image](https://github.com/hq0819/go-excels/assets/52909899/b4c498ed-4544-4ed5-a67b-250fc556316d)



使用excelTemplate指定列名即可

### 模板导出功能
创建excel模板文档 使用${fieldName}占位需要填写的数据，如果是切片使用${fe item:FieldName item.FieldName 去循环

![image](https://github.com/hq0819/go-excels/assets/52909899/9845fc2d-b8c0-402d-b577-742dd131b064)

```go

func main() {
	excel := excelTemplate.ReaderExcel[Person](`C:\Users\heqin11\Desktop\aaa.xlsx`, "")
	obj := Item{People: excel}
	excelTemplate.DoWrite("tmp.xlsx", "export.xlsx", obj)

}

type Item struct {
	People []Person
}

type Person struct {
	Name    string `excelTemplate:"姓名"`
	Age     int32 `excelTemplate:"年龄"`
	Address string `excelTemplate:"地址"`
}

```
![image](https://github.com/hq0819/go-excels/assets/52909899/21ee274d-6ccc-4205-868d-3eb1e0ddbea1)

