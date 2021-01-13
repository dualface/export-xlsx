  # 从 Excel 文件导出 JSON 文件

```
COPYRIGHT ALL RESERVED. (C) liaoyulei, https://github.com/dualface
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

## 导出配置

导出配置必须位于工作表的 A1 单元格，参考内容如下：

```
output: level_configs.json
index: level_id
header_row: 4
first_data_row: 5
```

- `output`: 指定输出的 JSON 文件
- `index`: 指定输出 JSON 时使用哪些字段进行索引
- `header_row`: 列头所在的行，定义了每一条数据包含哪些字段
- `first_data_row`: 数据的开始行

~

## 列头和数据

在工作表中，在 `header_row` 指定的行，从第一列开始定义列头。每一个列头定义数据里的一个字段。

从 `first_data_row` 指定的行开始，从第一列开始填写数据。每一行对应一条数据。

在 `header_row` 和 `first_data_row` 之间的行，可以添加各种备注信息，例如列头的说明，

示例（A - C 是列号，1 - 6 是行号）：

```
    #  A         |  B       |  C
    +------------+----------+------------
  1 |  output: level_configs.json
    |  index: level_id
    |  header_row: 4
    |  first_data_row: 5
    +------------+----------+------------
  2 |            |          |
    +------------+----------+------------
  3 |  配置每个级别的关卡会产生多少个 NPC
    +------------+----------+------------
  4 |  level_id  |  npc_id  |  quantity
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
    "level_id": "LEVEL_01",
    "npc_id": "NPC_01",
    "quantity": 100
}
```

所有数据都转换为字典后，再根据导出配置中 `index` 定义的索引，
提取每一个字典中特定的值作为 `KEY`，生成一个完整的字典。

每个 `KEY` 都是索引值，每个 `KEY` 对应每一条数据的字典。

因此最终输出的 JSON 是一个更大的字典：

```json
{
    "LEVEL_01": {
        "level_id": "LEVEL_01",
        "npc_id": "NPC_01",
        "quantity": 100
    },
    "LEVEL_02": {
        "level_id": "LEVEL_02",
        "npc_id": "NPC_02",
        "quantity": 100
    },
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
index: tech_id, level
header_row: 4
first_data_row: 5

tech_id  |  level  |  upgrade_cost
---------+---------+----------------
AAAAAA   |  1      |  100
AAAAAA   |  2      |  200
AAAAAA   |  3      |  300
BBBBBB   |  1      |  100
BBBBBB   |  2      |  200
BBBBBB   |  3      |  300
```

- `tech_id`: 科技 ID
- `level`: 科技的级别
- `upgrade_cost`: 升级的花费

第一级索引使用 `tech_id` 的值作为 `KEY`，第二级索引使用 `level`。生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "tech_id": "AAAAAA",
            "level": 1,
            "upgrade_cost": 100
        },
        "2": {
            "tech_id": "AAAAAA",
            "level": 2,
            "upgrade_cost": 200
        },
        "3": {
            "tech_id": "AAAAAA",
            "level": 3,
            "upgrade_cost": 300
        }
    },
    "BBBBBB": {
        "1": {
            "tech_id": "BBBBBB",
            "level": 1,
            "upgrade_cost": 100
        },
        "2": {
            "tech_id": "BBBBBB",
            "level": 2,
            "upgrade_cost": 200
        },
        "3": {
            "tech_id": "BBBBBB",
            "level": 3,
            "upgrade_cost": 300
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
index: tech_id, level
header_row: 4
first_data_row: 5

tech_id  |  level  |  upgrade_cost{  |  res_type  |  res_quantity  |  }
---------+---------+-----------------+------------+----------------+-----
AAAAAA   |  1      |  {              |  GOLD      |  100           |  }
AAAAAA   |  2      |  {              |  GOLD      |  200           |  }
AAAAAA   |  3      |  {              |  GOLD      |  300           |  }
```

生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "tech_id": "AAAAAA",
            "level": 1,
            "upgrade_cost": {
                "res_type": "GOLD",
                "res_quantity": 100
            }
        },
        "2": {
            "tech_id": "AAAAAA",
            "level": 2,
            "upgrade_cost": {
                "res_type": "GOLD",
                "res_quantity": 200
            }
        },
        "3": {
            "tech_id": "AAAAAA",
            "level": 3,
            "upgrade_cost": {
                "res_type": "GOLD",
                "res_quantity": 300
            }
        },
    }
}
```

定义嵌入字典的规则：

-   `upgrade_cost` 字段名后面增加了 `{` 符号，表示这个字段定义了一个嵌入的字典。

-   之后的 `}` 列头表示结束前一个嵌入字典的定义。

-   在 `upgrade_cost{` 和 `}` 之间定义的列头，就是嵌入字典的所有字段。
    -   上述示例中，`res_type` 和 `res_quantity` 就是嵌入字典的所有字段。

-   在填写数据时，每一条数据在字典开始的位置填写 `{`，在结束的位置填写 `}`。


~

## 包含多个字典的嵌入数组

前一个示例中，升级的花费只能有一种。现在修改为可以为多种：

```
output: technology_upgrade.json
index: tech_id, level
header_row: 4
first_data_row: 5

tech_id  |  level  |  upgrade_cost[  |  res_type  |  res_quantity  |  ]
---------+---------+-----------------+------------+----------------+-----
AAAAAA   |  1      |  {              |  GOLD      |  100           |  }
AAAAAA   |  2      |  {              |  GOLD      |  200           |
         |         |                 |  DIAMOND   |  20            |  }
AAAAAA   |  3      |  {              |  GOLD      |  300           |
         |         |                 |  DIAMOND   |  30            |
         |         |                 |  TICKET    |  3             |  }
```

生成的 JSON 如下：

```json
{
    "AAAAAA": {
        "1": {
            "tech_id": "AAAAAA",
            "level": 1,
            "upgrade_cost": [
                {
                    "res_type": "GOLD",
                    "res_quantity": 100
                }
            ]
        },
        "2": {
            "tech_id": "AAAAAA",
            "level": 2,
            "upgrade_cost": [
                {
                    "res_type": "GOLD",
                    "res_quantity": 200
                },
                {
                    "res_type": "DIAMOND",
                    "res_quantity": 20
                }
            ]
        },
        "3": {
            "tech_id": "AAAAAA",
            "level": 3,
            "upgrade_cost": [
                {
                    "res_type": "GOLD",
                    "res_quantity": 300
                },
                {
                    "res_type": "DIAMOND",
                    "res_quantity": 30
                },
                {
                    "res_type": "TICKET",
                    "res_quantity": 3
                }
            ]
        }
    }
}
```

定义嵌入数组的规则：

-   `upgrade_cost` 字段名后面增加了 `[` 符号，表示这个字段定义了一个嵌入的数组。

-   之后的 `]` 列头表示结束前一个嵌入数组的定义。

-   在 `upgrade_cost[` 和 `]` 之间定义的列头，就是嵌入数组里每个字典的所有字段。
    -   上述示例中，`res_type` 和 `res_quantity` 就是嵌入数组里每一个字典的所有字段。

-   在填写数据时，每一条数据在字典开始的位置填写 `{`，在结束的位置填写 `}`。

-   如果嵌入多个字典，就在第一个的字典开始位置填写 `{`，在最后一个字典的结束位置填写 `}`。
    -   示例中 `level: 1` 的数据就只定义了一个字典。
    -   示例中 `level: 2` 的数据定义了两个字典，`level: 3` 的数据定义了三个字典。


\-EOF\-
