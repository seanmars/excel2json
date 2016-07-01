# 2016/07/02

轉移到新的 repository, 之前的太雜亂且考慮到應該不需要使用到 web 相關套件,
所以開了一個新的 repository 來繼續撰寫此功能, 並且改名為 excel2json.
新的 repository 中只會有單純的 class 提供 API 來做 excel2json 的功能, 並且提供 cli 來使用.

# 2016/05/29

- Use Node.js:

  - Why:

    - 相對於 PHP, Python 比較可以馬上架構, 開發
    - 想練習 JavaScript

  - Pros:

    - 上手簡單
    - 有大量對此專案有用處的 plugin 可以應用
    - 可以承受大量的 Request

  - Cons:

    - Not enough diving in depth of JavaScript
    - 可能會有很多 nested callback, 維護上可能會有點複雜

- NPM
    - xlsx => https://github.com/SheetJS/js-xlsx
        - 需要讀取 excel 檔案, 眾多的套件中只有這個支援度較完整, 使用上也方便.
    - underscore => https://github.com/jashkenas/underscore
        - 最一開始需要使用這個套件是為了判斷是否為空的 Array.
    - jsonfile => https://github.com/jprichardson/node-jsonfile
        - 需要讀取及儲存 JSON 檔案, 此套件方便簡單使用, 且最多人下載.
    - commander => https://github.com/tj/commander.js/
        - 想要寫 cli, 剛好找到這個套件, 功能完整且最多人下載.
