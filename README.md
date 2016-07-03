# excel2json

Export the data with JSON format from excel;

## Features

- Import excel, Export JSON;
- Export JSON with prettify or minify;

## How To Use


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
var output = e2j.parse('data.xlsx', 'sheenname');
console.log(output);
```

## Todos

- Custom template;
- CLI;
