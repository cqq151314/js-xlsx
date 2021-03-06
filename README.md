# sheetjs 前端操作Excel的js插件

*一个excel可以导出多个sheet，可进行单元格的设置与合并*

## 内容目录

  1. js-xlsx安装
  2. 工作表
  3. sheets
  4. 单元格的合并
  5. 单元格的样式设置

## js-xlsx安装
   各种电子表格格式的解析器和编写器  

  - 在浏览器中使用

    ```jsx
    
    < script  src = “ https://unpkg.com/xlsx/dist/xlsx.full.min.js ” > < / script >
    
    ```

  - 用npm

    ```jsx
    $ npm install xlsx
    ```

- 样式(需要设置表格样式装)

    ```jsx
    npm install xlsx-style
    ```

## 工作表
  ```
      XXX| A  | B  | C  | D  | E  | F  | G  |
      ---+----+----+----+----+----+----+----+-
       1 | A1 | B1 | C1 | D1 | E1 | F1 | G1 |
       2 | A2 | B2 | C2 | D2 | E2 | F2 | G2 |
       3 | A3 | B3 | C3 | D3 | E3 | F3 | G3 |
```

## sheets
 - wb.SheetNames 工作簿中的工作表的有序列表。例如:['mySheet1', 'mySheet2', 'mySheet3']

 - wb.Sheets[sheetname] 返回表示工作表的对象。

  ## write()写入数据（导出write函数接受options参数）
 - type		------- [文件编码](https://github.com/SheetJS/js-xlsx#output-type)
 - bookType	------	工作簿类型（默认"xlsx“）


  ```jsx
    // 表格对象
    const wb = { SheetNames: ['Sheet1','Sheet2', 'Sheet3'], Sheets: {}};
    // write参数
    const wopts = { bookType: 'xlsx', type: 'binary' };
  
    //通过json_to_sheet转成单页(Sheet)数据,填充Sheets
    wb.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(data);
    wb.Sheets['Sheet2'] = XLSX.utils.json_to_sheet(data);
    wb.Sheets['Sheet3'] = XLSX.utils.json_to_sheet(data);
    // write方法写入数据
    const Blobs = new Blob([s2ab(XLSX.write(wb, wopts))], { type: "application/octet-stream" });
    const fileName = "这里是下载的文件名" + '.' + (wopts.bookType=="biff2"?"xls":wopts.bookType);
    saveAs(Blobs, fileName);

    // 字符串转字符流
    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
  ```

  ```jsx
    // 自定义的下载文件实现方式
    function saveAs(obj, fileName) { 
        var downloadEle = document.createElement('a');
        var objectUrl = URL.createObjectURL(obj);
        downloadEle.href = objectUrl; // URL对象创建
        downloadEle.download = fileName;
        document.body.appendChild(downloadEle);
        downloadEle.click();
        window.requestAnimationFrame(function(){
          document.body.removeChild(downloadEle);
          URL.revokeObjectURL(objectUrl); //用URL.revokeObjectURL()来释放这个object URL
        });
    }
  ```
- http://jsbin.com/nizawozaxu/6/edit?html,js,console,output


 ## Utility Functions（基于工作表）
- The sheet_to_* functions accept a worksheet and an optional options object.（XLSX.utils.sheet_add_json）

- The *_to_sheet functions accept a data object and an optional options object.（XLSX.utils.json_to_sheet）

## XLSX.utils.json_to_sheet的使用
- data: 必要参数
- header 制定表格的列顺序
  {header ： [ “ S ”，“ h ”，“ e ”，“ e_1 ”，“ t ”，“ J ”，“ S_1 ” ]}
- skipHeader: 如果为true，则不要在输出中包含标题行

## 表格对象
    1、每个不以!映射到单元格的键（使用A-1表示法）
    2、wb[address] 返回指定地址的单元格对象。

  - wb['!ref']：基于A-1的范围表示工作表范围。使用工作表的函数应使用此参数来确定范围。不处理在范围之外分配的单元格。
  - wb['!cols']： 列属性对象的数组。
  - wb['!rows']：行属性对象的数组。
  - wb['!merges']：与工作表中合并的单元格对应的范围对象数组。

  ## 合并单元格
  回顾官方文档，sheet提供了一个配置项!merges用来实现单元格合并,sheet['!merges']接受一个数组参数，数组对象的格式如下
  - wb.sheet['!merges']：与工作表中合并的单元格对应的范围对象数组。
  - c为列， r为行 从0开始

    ```jsx
     wb["!merges"] = [{
        s: { // s为开始
          c: C,
          r: R·
        },
        e: { // e为结束
          c: C,
          r: R }
    }]
    ``` 
    如下设置合并A1到D1
    ```jsx
    wb.sheet["!merges"] = [{
      s: { c: 1, r: 0 },
      e: { c: 4, r: 0 }
     }]
     ```

## 设置列宽
##### sheet提供了一个配置项!cols用来实现设置列宽

 - wb['!cols']=[{wpx: 100}, {wpx: 200}, {wpx: 300}, {wpx: 200}]

## 设置行高
##### sheet提供了一个配置项!rows用来实现设置列宽

 - wb['!rows']=[{hpx: 100}, {hpx: 200}, {hpx: 300}, {hpx: 200}]
    
## 样式设置（字体/背景颜色、对齐方式、边框）

  - 单元格样式有font、fill、alignment、和border。

    ```jsx
    wb["B1"].s = {
      font: { 
        sz: 14,
        bold: true,
        color: { rgb: "88FFAA99" }
        }, 
      fill: { 
        bgColor: { rgb: "88FF88" } 
        },
      alignment: {
        horizontal: "center" ,
        vertical: "center"
        }，
      border:{
        top:{
            style:'thick',
            color: { auto: 1}
        },
        left:{
            style:'thick',
            color: { auto: 1}
        },
        diagonal:{
            style:'thick',
            color: { rgb: "FFFFAA00" }
        },
        bottom:{
            style:'thick',
            color: { theme: "1", tint: "-0.1"},
        },
        right:{
            style:'thick',
            color: { indexed: 64}
        },
        diagonalUp:	true,
        diagonalDown: false
        }
      }
    ```
  ## 参考资料
  - [URL.createObjectURL和URL.revokeObjectURL](https://www.cnblogs.com/liulangmao/p/4262565.html)

  - [SheetJS/js-xlsx](https://github.com/SheetJS/js-xlsx#output-type)

  - [xlsx-style](https://www.npmjs.com/package/xlsx-style#cell-styles)

  - [Blob对象](https://www.cnblogs.com/hhhyaaon/p/5928152.html)
