  # 从 Excel 文件导出 JSON 文件

```
COPYRIGHT 2021 ALL RESERVED. (C) liaoyulei, https://github.com/dualface
```

这个工具的目标：

-   在策划定义的 Excel 文件中添加一些简单规则
-   从 Excel 文件生成特定格式的 JSON 文件

运行环境的要求：

-   Python 3.7+
-   使用 pip 安装 openpyxl 库

Excel 文件的格式要求：

-   必须保存为 `.xlsx` 格式
-   在需要导出的工作表中添加导出配置
-   在需要导出的工作表中添加列头和数据

~

## 使用方法

使用时，要转换的 Excel 文件需要按要求添加导出配置，然后使用命令行：

```bash
python3 export-xlsx.py test.xlsx
```

如果要一次性转换多个 Excel 文件，可以添加更多文件名到命令行中，或者使用通配符。例如：

```bash
python3 export-xlsx.py *.xlsx
```

~


## 导出配置

导出配置必须位于工作表的 A1 单元格，参考内容如下：

```
output: level_configs.json
index: levelId
header_row: 4
header_col: 1
first_data_row: 5
```

-   `output`: 指定输出的 JSON 文件
-   `index`: 指定输出 JSON 时使用哪些字段进行索引
-   `wrapper_field`: 将所有数据包装到为一个字典中的一个数组，`wrapper_field` 指定字段名，注意这个设置仅在没有设置 `index` 时生效
-   `header_row`: 列头所在的行，定义了每一条数据包含哪些字段
-   `header_col`: 列头所在的起始列，定义第一个字段从哪一列开始
-   `first_data_row`: 数据的开始行

~

## 列头和数据

在工作表中，在 `header_row` 指定的行，从 `header_col` 指定列开始定义列头。每一个列头定义数据里的一个字段。如果列头内容为空，则跳过这一列。

从 `first_data_row` 指定的行开始，从第一列开始填写数据。每一行对应一条数据。

在 `header_row` 和 `first_data_row` 之间的行，可以添加各种备注信息，例如列头的说明，

示例（A - C 是列号，1 - 6 是行号）：

```
    #  A         |  B       |  C
    +------------+----------+------------
  1 |  output: level_configs.json
    |  index: levelId
    |  header_row: 4
    |  header_col: 1
    |  first_data_row: 5
    +------------+----------+------------
  2 |            |          |
    +------------+----------+------------
  3 |  配置每个级别的关卡会产生多少个 NPC
    +------------+----------+------------
  4 |  levelId   |  NpcId   |  quantity
    +------------+----------+------------
  5 |  LEVEL_01  |  NPC_01  |  100
  6 |  LEVEL_02  |  NPC_02  |  100
```

-   `A1` 定义了导出配置
-   `A4` 到 `C4` 定义了三个列头
-   `A5:C5` 定义了第一行数据
-   `A6:C6` 定义了第二行数据

~

## 输出 JSON

工作表中，每一行数据，会按照列头指定的名字，构造成一个字典。

上面的示例中定义了三个列头，所以会生成如下的字典：

```json
"LEVEL_01": {
    "levelId": "LEVEL_01",
    "NpcId": "NPC_01",
    "quantity": 100
}
```

默认情况下，单元格中的值按照以下规则转换：

-   数字: 转换为 `number` 数字类型
-   空白, "null"（不区分大小写）: 转换为 `null` 空类型
-   "true" / "false"（不区分大小写）: 转换为 `boolean` 布尔类型
-   其他一律转换为 `string` 字符串类型

所有数据都转换为字典后，再根据导出配置中 `index` 定义的索引，
提取每一个字典中特定的值作为 `KEY`，生成一个完整的字典。

每个 `KEY` 都是索引值，每个 `KEY` 对应每一条数据的字典。

因此最终输出的 JSON 是一个更大的字典：

```json
{
    "LEVEL_01": {
        "levelId": "LEVEL_01",
        "NpcId": "NPC_01",
        "quantity": 100
    },
    "LEVEL_02": {
        "levelId": "LEVEL_02",
        "NpcId": "NPC_02",
        "quantity": 100
    }
}
```

~

如果要将所有数据包装为一个包含数组，那么可以用 `wrapper_field` 设置。

例如：

```
output: level_configs.json
header_row: 4
header_col: 1
first_data_row: 5
wrapper_field: array
```

输出的 JSON 格式将变成：

```json
{
    "array": [
        {
            "levelId": "LEVEL_01",
            "NpcId": "NPC_01",
            "quantity": 100
        },
        {
            "levelId": "LEVEL_02",
            "NpcId": "NPC_02",
            "quantity": 100
        }
    ]
}
```

> 注意：`wrapper_field` 仅在 `index` 没有设置时生效。

~


## 可选的列

如果某个列是可选的，那么可以在列头名字后面加上 `?` 符号。

对于可选的列头，如果行中该列没有填入内容或者填入了 `null`，则字典中不会包含该列。

示例：

```
  levelId   |  NpcId   |  quantity?
------------+----------+-------------
  LEVEL_01  |  NPC_01  |  null
  LEVEL_02  |  NPC_02  |  100
  LEVEL_03  |  NPC_03  |  null
```

输出：

```json
{
    "LEVEL_01": {
        "levelId": "LEVEL_01",
        "NpcId": "NPC_01"
    },
    "LEVEL_02": {
        "levelId": "LEVEL_02",
        "NpcId": "NPC_02",
        "quantity": 100
    },
    "LEVEL_03": {
        "levelId": "LEVEL_03",
        "NpcId": "NPC_03"
    }
}
```

~

## 两级索引

`index` 可以指定 1 个或 2 个字段名。使用 2 个字段名时，会生成包含两级索引的 JSON。

例如科技升级的表格中，每一个科技有多个级别，每个级别升级的花费都不同。
如果只使用科技的 ID 来做索引，就无法表达同一 ID 科技的多个级别了。

因此下面的示例使用了两级索引：

```
output: technology_upgrade.json
index: techID, level
header_row: 4
first_data_row: 5

techID  |  level  |  upgradeCost
--------+---------+---------------
AAAAAA  |  1      |  100
AAAAAA  |  2      |  200
AAAAAA  |  3      |  300
BBBBBB  |  1      |  100
BBBBBB  |  2      |  200
BBBBBB  |  3      |  300
```

-   `techID`: 科技 ID
-   `level`: 科技的级别
-   `upgradeCost`: 升级的花费

第一级索引使用 `techID` 的值作为 `KEY`，第二级索引使用 `level`。生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "techID": "AAAAAA",
            "level": 1,
            "upgradeCost": 100
        },
        "2": {
            "techID": "AAAAAA",
            "level": 2,
            "upgradeCost": 200
        },
        "3": {
            "techID": "AAAAAA",
            "level": 3,
            "upgradeCost": 300
        }
    },
    "BBBBBB": {
        "1": {
            "techID": "BBBBBB",
            "level": 1,
            "upgradeCost": 100
        },
        "2": {
            "techID": "BBBBBB",
            "level": 2,
            "upgradeCost": 200
        },
        "3": {
            "techID": "BBBBBB",
            "level": 3,
            "upgradeCost": 300
        }
    }
}
```

~

## 嵌入的字典

为了更好的组织数据结构，可以将一条记录里多个相关的字段定义为一个嵌入的字典。

示例如下：

```
output: technology_upgrade.json
index: techID, level
header_row: 4
first_data_row: 5

techID  |  level  |  upgradeCost{  |  resType  |  resQuantity  |  }
--------+---------+----------------+-----------+---------------+-----
AAAAAA  |  1      |             {  |  GOLD     |  100          |  }
AAAAAA  |  2      |             {  |  GOLD     |  200          |  }
AAAAAA  |  3      |             {  |  GOLD     |  300          |  }
```

生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "techID": "AAAAAA",
            "level": 1,
            "upgradeCost": {
                "resType": "GOLD",
                "resQuantity": 100
            }
        },
        "2": {
            "techID": "AAAAAA",
            "level": 2,
            "upgradeCost": {
                "resType": "GOLD",
                "resQuantity": 200
            }
        },
        "3": {
            "techID": "AAAAAA",
            "level": 3,
            "upgradeCost": {
                "resType": "GOLD",
                "resQuantity": 300
            }
        }
    }
}
```

定义嵌入字典的规则：

-   `upgradeCost` 字段名后面增加了 `{` 符号，表示这个字段定义了一个嵌入的字典。

-   之后的 `}` 列头表示结束前一个嵌入字典的定义。

-   在 `upgradeCost{` 和 `}` 之间定义的列头，就是嵌入字典的所有字段。
    -   上述示例中，`resType` 和 `resQuantity` 就是嵌入字典的所有字段。

-   在填写数据时，每一条数据在字典开始的位置填写 `{`，在结束的位置填写 `}`。

-   如果字典列头为可选，那么空字典将不会被包含在结果中。

~

## 包含多个字典的嵌入数组

前一个示例中，升级的花费只能有一种。现在修改为可以为多种：

```
output: technology_upgrade.json
index: techID, level
header_row: 4
first_data_row: 5

techID  |  level  |  upgradeCost[  |  resType  |  resQuantity  |  ]
--------+---------+----------------+-----------+---------------+-----
AAAAAA  |  1      |             {  |  GOLD     |  100          |  }
AAAAAA  |  2      |             {  |  GOLD     |  200          |
        |         |                |  DIAMOND  |  20           |  }
AAAAAA  |  3      |             {  |  GOLD     |  300          |
        |         |                |  DIAMOND  |  30           |
        |         |                |  TICKET   |  3            |  }
```

生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "techID": "AAAAAA",
            "level": 1,
            "upgradeCost": [
                {
                    "resType": "GOLD",
                    "resQuantity": 100
                }
            ]
        },
        "2": {
            "techID": "AAAAAA",
            "level": 2,
            "upgradeCost": [
                {
                    "resType": "GOLD",
                    "resQuantity": 200
                },
                {
                    "resType": "DIAMOND",
                    "resQuantity": 20
                }
            ]
        },
        "3": {
            "techID": "AAAAAA",
            "level": 3,
            "upgradeCost": [
                {
                    "resType": "GOLD",
                    "resQuantity": 300
                },
                {
                    "resType": "DIAMOND",
                    "resQuantity": 30
                },
                {
                    "resType": "TICKET",
                    "resQuantity": 3
                }
            ]
        }
    }
}
```

定义嵌入数组的规则：

-   `upgradeCost` 字段名后面增加了 `[` 符号，表示这个字段定义了一个嵌入的数组。

-   之后的 `]` 列头表示结束前一个嵌入数组的定义。

-   在 `upgradeCost[` 和 `]` 之间定义的列头，就是嵌入数组里每个字典的所有字段。
    -   上述示例中，`resType` 和 `resQuantity` 就是嵌入数组里每一个字典的所有字段。

-   在填写数据时，每一条数据在字典开始的位置填写 `{`，在结束的位置填写 `}`。

-   如果嵌入多个字典，就在第一个的字典开始位置填写 `{`，在最后一个字典的结束位置填写 `}`。
    -   示例中 `level: 1` 的数据就只定义了一个字典。
    -   示例中 `level: 2` 的数据定义了两个字典，`level: 3` 的数据定义了三个字典。

-   空字典不会被包含在数组中。

-   如果数组列头为可选，那么空数组将不会被包含在结果中。

~

## 不含字段名的匿名数组

如果记录中的数组只是需要保存不带字段名的值，那么可以将字段定义为匿名数组。

匿名数组以 `#` 开头。该数组中所有字段值都会忽略掉字段名，字段值则会依次添加到数组中。

```
output: technology_upgrade.json
index: techID, level
header_row: 4
first_data_row: 5

techID  |  level  |  #upgradeCost[  |  resType  |  resQuantity  |  ]
--------+---------+----------------+-----------+---------------+-----
AAAAAA  |  1      |              {  |  GOLD     |  100          |  }
AAAAAA  |  2      |              {  |  GOLD     |  200          |
        |         |                 |  DIAMOND  |  20           |  }
AAAAAA  |  3      |              {  |  GOLD     |  300          |
        |         |                 |  DIAMOND  |  30           |
        |         |                 |  TICKET   |  3            |  }
```


生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "techID": "AAAAAA",
            "level": 1,
            "upgradeCost": [
                "GOLD",
                100
            ]
        },
        "2": {
            "techID": "AAAAAA",
            "level": 2,
            "upgradeCost": [
                "GOLD",
                200,
                "DIAMOND",
                20
            ]
        },
        "3": {

"techID": "AAAAAA",
            "level": 3,
            "upgradeCost": [
                "GOLD",
                300,
                "DIAMOND",
                30,
                "TICKET",
                3
            ]
        }
    }
}
```

定义匿名数组的规则：

-   空白值或者 `null` 不会被包含在数组中。

-   如果数组列头为可选，那么空数组将不会被包含在结果中。

~


## 为字段指定类型

默认情况下，会根据字段值转换为适当的类型。如果有需要，可以在字段名后面添加字段类型。字段名和类型之间用 ":" 分割。

例如：

```
levelId   |  NpcId   |  quantity:string
----------+----------+-------------------
LEVEL_01  |  NPC_01  |  100
```

为 `quantity` 指定类型后，输出的 JSON 中，该字段的值就是字符串类型：

```json
"LEVEL_01": {
    "levelId": "LEVEL_01",
    "NpcId": "NPC_01",
    "quantity": "100"
}
```

目前支持的类型有：

- `auto`: 自动判断，等同于不指定类型
- `string`: 转换为字符串
- `int`: 转换为整数
- `float`: 转换为浮点数
- `bool`: 转换为 boolean 类型
- `vec2`: 输入格式是 x,y，输出为一个字典 `{"x": x, "y": y}`
- `vec2int`: 输出格式同上，但转换为整数

例如：

```
deviceId   |  size:vec2
-----------+-------------
DEVICE_01  |  10, 20
```

输出为：

```json
"DEVICE_01": {
    "deviceId": "DEVICE_01",
    "size": {
        "x": 10.0,
        "y": 20.0
    }
}
```

\-EOF\-
