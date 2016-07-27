# excel-to-json-template

Export the data with JSON format from excel;

- [Features](#features)
- [How To Use](#howtouse)
- [Todos](#todos)
- [Change logs](#logs)

## <a name="features"></a>Features

- Import excel, Export JSON;
- Export JSON with prettify or minify;
- Custom JSON template;

## <a name="howtouse"></a>How To Use

### Excel
---
Tags:

- \! => Ignore
    - Column == A, it will ignore all the same column and row;
    - Column != A, just ignore the same column;
- \# => Title
    - The title of data in excel, it will transform to property's name of Object;

### Example
---

- Use code:
```javascript
var e2jt = require('./index.js');
e2jt.loadTemplate('path/template.json', function(err, jsonObj) {
    var data = e2jt.parse('path/excel.xlsx', 'sheet name', jsonObj);
    e2jt.save('path/to/output/data.json', data);
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

### 2016/07/28
---
FEATURE

- 增加 tag: attribute(^), 讓使用者可以在 JSON 的 top-level 增加 attribute;
    - tag attribute 的位子一定要在 tag title 上方
    - tag attribute 由同 row 的三個連續的 column 組成, 順序為 tag > property name > value

IMPROVE

- 修改取得資料的迴圈只需要跑 tag title 以下的 row

OTHER

- 引入 mochajs 增加 Unit test;

### 2016/07/09
---
- Refactoring some code;(rename, add some check...)

### 2016/07/09
---
- Add CLI;

### 2016/07/08
---
- Update the parse function, now can input empyt or null template, it will direct parse(no any transform) all the title in excel to JSON;

### 2016/07/06
---
- Add feature custom template;
