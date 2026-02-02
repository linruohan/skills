---
name: pandas-xlsx
description: "使用 pandas 库快速创建和操作 Excel 文件的专业技能。支持创建空白表格、添加数据、设置格式、读取分析 Excel 文件等功能。适用于数据处理、报表生成、数据分析等场景。"
license: MIT
---

# Pandas Excel 操作工具

## 概述

这是一个专业的 Excel 文件操作技能，使用 pandas 库提供快速、高效的 Excel 文件创建、编辑和分析解决方案。适用于数据处理、自动化报表生成、数据分析等场景。

## 前置条件

确保已安装必要的依赖库：

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

## 高级 Excel 操作

### 1. 创建复杂表格结构

```python
import pandas as pd
from datetime import datetime

# 创建数据字典
data = {
    '姓名': ['张三', '李四', '王五'],
    '年龄': [25, 30, 35],
    '部门': ['技术部', '市场部', '财务部'],
    '入职日期': [datetime(2020, 1, 1), datetime(2019, 5, 15), datetime(2021, 3, 20)],
    '薪资': [8000, 9500, 11000]
}

df = pd.DataFrame(data)
df.to_excel('员工信息表.xlsx', index=False, sheet_name='员工信息')
```

### 2. 创建多工作表 Excel

```python
import pandas as pd

# 创建第一个工作表
data1 = {
    '产品名称': ['产品A', '产品B', '产品C'],
    '销量': [100, 150, 200],
    '收入': [50000, 75000, 100000]
}
df1 = pd.DataFrame(data1)

# 创建第二个工作表
data2 = {
    '月份': ['1月', '2月', '3月'],
    '成本': [30000, 35000, 40000]
}
df2 = pd.DataFrame(data2)

# 写入多工作表 Excel
with pd.ExcelWriter('年度报表.xlsx', engine='openpyxl') as writer:
    df1.to_excel(writer, sheet_name='销售报表', index=False)
    df2.to_excel(writer, sheet_name='成本报表', index=False)
```

### 3. 读取和分析 Excel 文件

```python
import pandas as pd

# 读取 Excel 文件
df = pd.read_excel('员工信息表.xlsx', sheet_name='员工信息')

# 数据分析
print("数据概览:")
print(df.describe())

print("\n部门统计:")
print(df['部门'].value_counts())

print("\n平均薪资:")
print(df['薪资'].mean())
```

### 4. 数据清洗和处理

```python
import pandas as pd

# 创建包含缺失值的示例数据
data = {
    '姓名': ['张三', '李四', None, '王五'],
    '年龄': [25, 30, 35, None],
    '薪资': [8000, None, 11000, 9000]
}
df = pd.DataFrame(data)

# 处理缺失值
df_filled = df.fillna({
    '姓名': '未知',
    '年龄': df['年龄'].mean(),
    '薪资': df['薪资'].mean()
})

# 保存处理后的数据
df_filled.to_excel('清洗后的数据.xlsx', index=False)
```

## 关键参数说明

### to_excel 参数
- `index=False`：不写入行索引
- `engine='openpyxl'`：指定使用 openpyxl 引擎
- `sheet_name`：工作表名称
- `startrow`：写入起始行
- `startcol`：写入起始列

### read_excel 参数
- `sheet_name`：指定读取的工作表名称
- `header=0`：指定行作为列名
- `index_col`：指定列作为索引
- `dtype`：指定列的数据类型

## 使用场景

### 1. 批量报表生成

```python
import pandas as pd
import os

# 批量生成报表
def generate_monthly_reports(year, month):
    base_data = {
        '月份': [f'{year}-{month:02d}-01'],
        '销售额': [50000],
        '成本': [30000],
        '利润': [20000]
    }
    
    df = pd.DataFrame(base_data)
    filename = f'{year}_{month:02d}_月度报表.xlsx'
    df.to_excel(filename, index=False)
    return filename

# 生成全年报表
for month in range(1, 13):
    generate_monthly_reports(2024, month)
```

### 2. 数据导出

```python
import pandas as pd

# 假设从数据库或其他数据源获取数据
data = pd.read_sql('SELECT * FROM sales_data', connection)

# 导出为 Excel
data.to_excel('销售数据导出.xlsx', 
               index=False,
               sheet_name='原始数据',
               startrow=1,  # 从第二行开始写
               startcol=1)  # 从第二列开始写

# 添加汇总工作表
summary = data.groupby('product')['amount'].sum().reset_index()
with pd.ExcelWriter('销售数据导出.xlsx', engine='openpyxl', mode='a') as writer:
    summary.to_excel(writer, sheet_name='汇总统计', index=False)
```

### 3. 数据验证和质量检查

```python
import pandas as pd

def validate_excel_data(filepath):
    """验证 Excel 数据质量"""
    df = pd.read_excel(filepath)
    
    # 检查缺失值
    missing_values = df.isnull().sum()
    
    # 检查数据类型
    data_types = df.dtypes
    
    # 检查重复行
    duplicates = df.duplicated().sum()
    
    # 创建质量报告
    quality_report = pd.DataFrame({
        '检查项目': ['缺失值数量', '重复行数量', '总行数', '总列数'],
        '数值': [
            missing_values.sum(),
            duplicates,
            len(df),
            len(df.columns)
        ]
    })
    
    return quality_report

# 使用示例
quality_report = validate_excel_data('员工信息表.xlsx')
quality_report.to_excel('数据质量报告.xlsx', index=False)
```

## 最佳实践

### 1. 内存管理

```python
import pandas as pd
import gc

# 处理大型 Excel 文件时使用分块读取
chunk_size = 10000
chunks = pd.read_excel('large_file.xlsx', chunksize=chunk_size)

for chunk in chunks:
    # 处理每个数据块
    process_chunk(chunk)
    
    # 及时释放内存
    del chunk
    gc.collect()
```

### 2. 错误处理

```python
import pandas as pd
import os

def safe_excel_operation(filename, operation='save', data=None):
    """安全的 Excel 操作"""
    try:
        if operation == 'save' and data is not None:
            # 检查目录是否存在
            os.makedirs(os.path.dirname(filename), exist_ok=True)
            data.to_excel(filename, index=False)
            print(f"文件保存成功: {filename}")
            return True
        elif operation == 'read':
            if os.path.exists(filename):
                df = pd.read_excel(filename)
                return df
            else:
                print(f"文件不存在: {filename}")
                return None
        else:
            print("不支持的操作类型")
            return False
    except Exception as e:
        print(f"Excel 操作失败: {e}")
        return False

# 使用示例
df = safe_excel_operation('data.xlsx', operation='read')
if df is not None:
    safe_excel_operation('processed_data.xlsx', operation='save', data=df)
```

### 3. 性能优化

```python
import pandas as pd

# 批量处理优化
def batch_process_excel_files(file_list):
    """批量处理 Excel 文件"""
    results = []
    
    for file in file_list:
        # 读取文件
        df = pd.read_excel(file)
        
        # 数据处理
        processed = df.groupby('category')['value'].sum().reset_index()
        
        # 保存结果
        output_file = f'processed_{file}'
        processed.to_excel(output_file, index=False)
        
        results.append({
            '原文件': file,
            '处理后文件': output_file,
            '数据行数': len(df)
        })
    
    # 生成处理报告
    report = pd.DataFrame(results)
    report.to_excel('批量处理报告.xlsx', index=False)
    return report
```

## 数据格式支持

### 1. 支持的数据类型
- **数值型**：int, float, decimal
- **文本型**：string, object
- **日期时间**：datetime, timestamp
- **布尔型**：bool
- **分类数据**：category

### 2. 数据转换示例

```python
import pandas as pd

# 创建包含各种数据类型的数据
data = {
    '数值列': [1, 2, 3, 4, 5],
    '文本列': ['A', 'B', 'C', 'D', 'E'],
    '日期列': pd.date_range('2024-01-01', periods=5),
    '布尔列': [True, False, True, False, True],
    '分类列': ['低', '中', '高', '中', '低']
}

df = pd.DataFrame(data)

# 设置数据类型
df['分类列'] = df['分类列'].astype('category')
df['布尔列'] = df['布尔列'].astype('bool')

# 保存为 Excel
df.to_excel('多类型数据.xlsx', index=False)
```

## 总结

- **核心逻辑**：创建 DataFrame → 数据处理 → 保存为 Excel
- **命令行执行**：使用 `python -c "import pandas as pd; pd.DataFrame(...).to_excel(...)"` 进行单行命令执行
- **必装依赖**：`pandas` 和 `openpyxl` 是必需的核心库
- **适用场景**：适合数据处理、报表生成、数据分析等场景
- **扩展性**：支持复杂的 Excel 操作，包括多工作表、格式设置、数据验证等

## 注意事项

1. 确保已安装 pandas 和 openpyxl 库
2. Excel 版本：生成的文件兼容 Excel 2007+ (.xlsx 格式)
3. 内存管理：处理大型文件时注意内存使用
4. 编码问题：建议使用 UTF-8 编码
5. 性能考虑：对于大数据集，考虑分块处理或使用更高效的格式