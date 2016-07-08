# excel2json

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
```javascript
var e2j = require('./excel2json.js');
e2j.load('path/template.json', function(err, jsonObj) {
    var data = e2j.parse('path/excel.xlsx', 'sheet name', jsonObj);
    e2j.save('path/to/output/data.json', data);
});
```

## <a name="todos"></a>Todos

- CLI;

## <a name="logs"></a>Change logs

### 2016/07/06
---
- Add feature custom template;
