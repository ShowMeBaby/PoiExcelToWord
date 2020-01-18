# PoiExcelToWord

> 一个基于Apache POI4.1.1开发,从Excel中读取变量与表格,并填充到Word中的命令行工具

**环境要求: JAVA1.8**

使用介绍:

1. 进入guide文件夹

2. 于命令行输入
```shell
   java -DxlsxPath='test.xlsx' -DinputPath='test.docx' -DoutPath='result.docx' -jar ExcelToWord.jar
```
3. 其中xlsxPath为 数据来源路径,  inputPath为word模版路径, outPath为解析后导出的文件路径

4. guide下的ExcelToWord.jar为打包好后的包, 默认输出json, 可以直接拿去通过命令行调用

#### 注意事项

1. 所有word模版中的变量格式均为**${变量名}**格式,表格也是
2. excel文档中需要有一个名为**变量表**的,A列表为变量名,B列为变量值,  支持引用, 表格的引用格式为 **"工作表名称  A1:510"**
3. 详情可参考test.xlsx
#### License

PoiExcelToWord is released under the[Apache 2.0 license](https://www.apache.org/licenses/LICENSE-2.0.html);