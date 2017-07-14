# Excel_Handler API Doc

## 1. 格式清洗

将传入Excel文件的标题行合并成一行，返回xlsx格式的文件。  

URL : http://localhost/api/clean_xl

Method : POST  

参数 :  

| 参数名称     | 描述                 | 必须   | 唯一/多个 |
| -------- | ------------------ | ---- | ----- |
| filename | 传入Excel文件的文件名（含后缀） | yes  | 唯一    |

返回值 :
| 名称          | 描述                                       | 说明     |
| ----------- | ---------------------------------------- | ------ |
| errno       | 执行结果错误代码                                 |        |
| msg         | 执行结果错误信息                                 |        |
| path2file   | 清洗后的Excel文件的存储路径                         |        |
| type_scheme | Excel文件各数据表（Sheet）的各列名（字段名）及对应数据类型名的数据字典 | json格式 |
