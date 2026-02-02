# Pandas Excel 创建工具

使用 pandas 创建 Excel 文件的快速方法。

## 前置条件

确保已安装 pandas 和 openpyxl：

```bash
pip install pandas openpyxl
```

## 快速创建 Excel

### 创建空的 Excel 文件

```bash
python -c "import pandas as pd; pd.DataFrame().to_excel('empty_file.xlsx', index=False, engine='openpyxl')"
```

### 创建只有表头的 Excel

```bash
python -c "import pandas as pd; pd.DataFrame(columns=['列1','列2','列3']).to_excel('header_only.xlsx', index=False, engine='openpyxl')"
```

### 创建带数据的 Excel

```bash
python -c "import pandas as pd; pd.DataFrame({'列1':[1,2,3],'列2':['a','b','c'],'列3':[10,20,30]}).to_excel('data_file.xlsx', index=False, engine='openpyxl')"
```

## 使用说明

- `index=False`：不写入行索引
- `engine='openpyxl'`：指定使用 openpyxl 引擎
- 数据格式：字典形式，键为列名，值为列表数据

## 示例

创建一个包含数字 1,2,3 的 Excel：

```bash
python -c "import pandas as pd; pd.DataFrame({'数据':[1,2,3]}).to_excel('test.xlsx', index=False, engine='openpyxl')"
```
