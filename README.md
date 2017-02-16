# excel-to-json-template

Export the data with JSON format from excel;

- [Features](#features)
- [Usage](#usage)
- [Todos](#todos)
- [Change logs](#logs)

## <a name="features"></a>Features

- Import excel, Export JSON;
- Export JSON with prettify or minify;
- Custom JSON template;

## <a name="usage"></a>Usage

### Excel
---
Tag:

- \! => Ignore
    - Column == A, it will ignore all the same column and row;
    - Column != A, just ignore the same column;
- \# => Title
    - The title of data in excel, it will transform to property's name of Object;
- ^ => Top-level attributes
    - If use dict(--dict true), all the top-level attributes will be ignored.

| A | B | C |
|---|---|---|
| ^ | key | value |

- \~ => The name of key.

| A | B |
|---|---|
| ~ | name of key |

### Template

- \$ => Use $title_name can map the value to be key in json;

### Example
---

- Use code:
```javascript
var e2jt = require('./index.js');
e2jt.loadTemplate('./template/file/path/file.json', function (err, jsonObj) {
    var data = e2jt.parse('./data/file/path/data.xlsx', 'sheet_name', jsonObj);
    e2jt.save('path/to/output/output.json', data);
});
```

- Use cli:
```
e2jt /path/of/file
```
More about cli:
```
e2jt -h
```

## <a name="todos"></a>Todos

- CLI;
    - Use command to import file(list of file) and parse all of then;

## <a name="logs"></a>Change logs

### 2017-02-17
---
FEATURE

- 增加 CLI 選項，是否自動加入屬性 sheetname。--sheetname [true|false] default = false。
- 增加 CLI 選項，輸出是否為 key-value。--dict [true|false] default = false。 NOTE: 設為 true 時，所有 Top-level attributes(^) 將會被忽略。
- 增加 CLI 選項，資料列的 Key 名稱。--key <name of key>。
- 增加 Tag:"~"，資料表中 Tag:"~" 欄位右邊的值會編譯成資料列的 Key 名稱。 NOTE: 此屬性的優先順序比 CLI 的 --key 低。
- 資料列的 Key 名稱預設為 datas。

### 2017-02-16
---
FEATURE

- 增加 $ 前導符號應用，在 Template 當中使用 ${title_name} 將可把數值編譯成 JSON 的 KEY。

Template:

```JSON
{
    "$id": {
        "name": "name",
        "age": "age"
    }
}
```

Output:

```JSON
{
    "0001": {
        "name": "Foobar",
        "age": "23"
    }
}
```

### 2016-09-06
---
FIXED

- Now can use the recursively Object, Array in template file;

OTHER

- Add more test case;

### 2016-07-28
---
FEATURE

- 增加 tag: attribute(^), 讓使用者可以在 JSON 的 top-level 增加 attribute;
    - tag attribute 的位子一定要在 tag title 上方
    - tag attribute 由同 row 的三個連續的 column 組成, 順序為 tag > property name > value

IMPROVE

- 修改取得資料的迴圈只需要跑 tag title 以下的 row

OTHER

- 引入 mochajs 增加 Unit test;

### 2016-07-09
---
- Refactoring some code;(rename, add some check...)

### 2016-07-09
---
- Add CLI;

### 2016-07-08
---
- Update the parse function, now can input empyt or null template, it will direct parse(no any transform) all the title in excel to JSON;

### 2016-07-06
---
- Add feature custom template;
