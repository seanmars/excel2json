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

- CLI
    - Use command to import file(list of file) and parse all of then.

## <a name="logs"></a>Change logs

### 2016/07/09
---
- Add CLI;

### 2016/07/08
---
- Update the parse function, now can input empyt or null template, it will direct parse(no any transform) all the title in excel to JSON;

### 2016/07/06
---
- Add feature custom template;
