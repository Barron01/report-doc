# report-doc
# 批量生成 DOC 文档的宏说明

## 概述
此宏旨在帮助用户批量生成 DOC 文档，提高工作效率。用户只需提供数据源，宏将自动创建并填充文档内容。

## 需求
- Microsoft Word
- 数据源文件（如 Excel 或 CSV）

## 步骤

### 1. 准备数据源
确保数据源文件包含所需的信息。每行应代表一个文档，每列代表一个字段。例如：

| REPORTDATE     | PROBLEMHANDLING        | TECHNICALADVISORYSERVICE | WORKTHISWEEK                                                 | SYSTEMRUNNINGSTATUSMONITORINGRESULT                          | ACTIVESYSTEMMAINTENANCE                                      | PROBLEMSUMMARY                                               |
| -------------- | ---------------------- | ------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ | ------------------------------------------------------------ |
| 2023年10月11日 | 本周新增0个，关闭0个。 | 有1人次                  | 本周自2023年10月5日到2023年10月11日止，总计7天，其中工作日5天,系统每天使用时间08时30分-18时00分。 | CPU状态：使用率最大值66.67%、平均值35.32%、正常；^p 内存状态：使用率最大值64.72%、平均值46.88%、正常；^p 存储状态：本周新增数据量2057 MB、增长率0.41 %、当前存储使用率50.41%、正常；^p 网络状态：吞吐量最大值10698.85KB/s、最小值67.23KB/s、平均值1057.77KB/s、正常；^p 虚机资源：虚机数量15、在线量15、在线率100%、正常；^p 用户状态：监控周期内总访问量32673次、最大并发量103；^p 告警状态：紧急告警0条，关闭0条；主要告警0条，关闭0条；一般告警0条，关闭0条。^p 新增数据：监控周期内新增加数据0条、增长率0% | 系统巡检：对系统共主动巡检7次，巡检发现问题0个，解决0个；^p 系统优化：对系统优化0次；^p 系统升级：完成系统升级0次； | 本周内共收集问题0条，其中高风险问题0条、中风险问题0条、低风险问题0条；^p 本周内问题已关闭0条、关闭率0%；处理中问题0条、待处理问题0条；^p 截止到本周共计收集问题0条、其中高风险问题0条、中风险问题0条、低风险问题0条；^p 截止到本周已关闭0条、关闭率0%；处理中问题0条、待处理问题0条； |

### 2. 打开 Word 并启用开发工具
在 Microsoft Word 中，确保启用开发工具选项卡：
1. 打开 Word。
2. 进入“文件” > “选项” > “自定义功能区”。
3. 勾选“开发工具”选项卡，点击“确定”。

### 3. 创建宏
1. 在“开发工具”选项卡中，点击“宏”按钮。
2. 输入宏的名称，如 `BatchCreateDocs`，点击“创建”。
3. 在弹出的 VBA 编辑器中，输入以下代码：

    ```vba
    Sub BatchReplaceAndSave()
        Dim originalDoc As Document
        Dim newDoc As Document
        Dim findTexts() As String
        Dim replaceTexts() As String
        Dim savePath As String
        Dim excelApp As Object
        Dim excelWorkbook As Object
        Dim excelWorksheet As Object
        Dim colNum As Integer
        Dim newRow As Integer
        Dim col As Integer
        Dim i As Integer
        Dim maxLength As Integer
        Dim segmentLength As Integer
        Dim segmentCount As Integer
        Dim segmentIndex As Integer
        Dim searchText As String
        Dim replaceString As String
        
        ' 打开 Excel 文件
        Set excelApp = CreateObject("Excel.Application")
        Set excelWorkbook = excelApp.Workbooks.Open("C:\Users\Acho\Desktop\文档变量.xlsx")
        Set excelWorksheet = excelWorkbook.Sheets(1)
        
        ' 获取 Excel 表格中的列数
        colNum = excelWorksheet.Cells(1, excelWorksheet.Columns.Count).End(-4159).Column
        
        ' 如果列数为 0，则说明表格为空，给出错误提示并退出宏
        If colNum = 0 Then
            MsgBox "Excel 表格为空，请检查。"
            Exit Sub
        End If
        
        ' 设置保存路径
        savePath = "C:\Users\Acho\Desktop\"
        
        ' 初始化替换的字符串数组
        ReDim replaceTexts(1 To colNum)
        
        ' 从 Excel 表格中读取第一行作为替换的字符串
        For col = 1 To colNum
            replaceTexts(col) = excelWorksheet.Cells(1, col).Value
        Next col
        
        ' 循环遍历每一行数据（除第一行）
        For newRow = 2 To excelWorksheet.Cells(excelWorksheet.Rows.Count, 1).End(-4162).Row
        
            Dim fileName As String
            fileName = Trim(excelWorksheet.Cells(newRow, 1).Value) ' 使用 Trim 函数去除可能的前后空格
    
            
            ' 打开原始文档
            Set originalDoc = Documents.Open("C:\Users\试运行工作周报模板.docx")
            
            ' 创建新文档
            Set newDoc = Documents.Add
            
            ' 复制原始文档内容到新文档
            originalDoc.Content.Copy
            newDoc.Content.Paste
            
            ' 关闭原始文档
            originalDoc.Close False
            
            ' 循环遍历每个查找和替换字符串
            For i = 1 To colNum
                ' 获取要查找和替换的文本
                searchText = replaceTexts(i)
                replaceString = excelWorksheet.Cells(newRow, i).Value
                
                ' 计算替换文本的长度
                maxLength = Len(replaceString)
                ' 设置每段文本的长度
                segmentLength = 200
                ' 计算分段数量
                segmentCount = (maxLength \ segmentLength) + IIf(maxLength Mod segmentLength > 0, 1, 0)
                
                ' 分段替换
                For segmentIndex = 1 To segmentCount
                    ' 获取当前段的文本
                    Dim startChar As Integer
                    startChar = (segmentIndex - 1) * segmentLength + 1
                    Dim endChar As Integer
                    endChar = startChar + segmentLength - 1
                    If endChar > maxLength Then
                        endChar = maxLength
                    End If
                    Dim segmentReplaceString As String
                    segmentReplaceString = Mid(replaceString, startChar, endChar - startChar + 1)
    
                    ' 替换字符串
                    newDoc.Content.Find.Execute findText:=searchText, ReplaceWith:=segmentReplaceString & searchText, Replace:=wdReplaceAll
                Next segmentIndex
            Next i
            
            ' 保存新文档
            
            ' 检查文件名是否为空
            If fileName <> "" Then
                ' 文件名不为空时执行保存操作
                fileName = Replace(fileName, "/", "-") ' 替换文件名中的斜杠
                fileName = Replace(fileName, ":", "-") ' 替换文件名中的冒号
                newDoc.SaveAs2 fileName:=savePath & fileName & ".docx", FileFormat:=wdFormatXMLDocument
            Else
                ' 文件名为空时给出提示或执行其他操作
                MsgBox "第 " & newRow & " 行的文件名为空，跳过保存。"
            End If
                      
            ' 关闭新文档
            newDoc.Close
        Next newRow
        
        ' 关闭 Excel 文件
        excelWorkbook.Close False
        excelApp.Quit
        Set excelWorksheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
    End Sub
    ```

### 4. 运行宏
1. 返回 Word。
2. 在“开发工具”选项卡中，点击“宏”按钮。
3. 选择 `BatchCreateDocs`，点击“运行”。

### 5.示例

​	1、template目录中有示例文件，可以直接执行。

## 注意事项

- 确保数据源文件路径正确。
- 确保保存文档的路径存在。
- 根据实际需求修改宏中的字段和路径。

通过以上步骤，您可以轻松批量生成 DOC 文档。希望此宏能够提升您的工作效率！