# js-xlsx/导出报表进阶

*一个excel可以导出多个sheet，可进行单元格的设置与合并*

## 内容目录

  1. a标签的download导出excel
  1. js-xlsx介绍及使用
  1. sheets
  1. 单元格的合并
  1. 单元格的样式设置
  1. demo

## <a>标签

  - <a>标签的download属性实现点击下载

    ```jsx
    // 如果没传maps则取数据的字段作为maps { 姓名: 'name', 年龄: 'age' }、
    const keys = Object.keys(maps);

    const csvStr = BOM + [
      keys.map(key => maps[key]).toString(),
      ...dataSource.map(item => keys.map(key => item[key]).toString())
    ].join('\n');

    const downloadEle = document.createElement('a');

    downloadEle.href = `data:attachment/csv,${encodeURI(csvStr)}`;
    downloadEle.target = '_blank';
    downloadEle.download = fileName;

    document.body.appendChild(downloadEle);
    downloadEle.click();
    document.body.removeChild(downloadEle);
    }
    ```

  - csv文件分行用 “,”，而分列用\n无效，必须使用encodeURI进行编码.
  - 导出csv 格式， 使用Excel 打开会发现中文是乱码，但是用其他文本程序打开确是正常的,原因就是少了一个 BOM头 。  \ufeff。

## js-xlsx介绍
   各种电子表格格式的解析器和编写器  

  - 在浏览器中使用

    ```jsx
    
    < script  src = “ https://unpkg.com/xlsx/dist/xlsx.full.min.js ” > < / script >
    
    ```

  - 用npm

    ```jsx
    1、$ npm install xlsx
    2、添加脚本标记 < script  lang = “ javascript ”  src = “ dist / xlsx.full.min.js ” > < / script >
    ```

## 单元格
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

   ```jsx
   var dataSource = [
       {
         "id": 1, "name": "小明", "age": 22,
       }
   ]

    var tmpdata = dataSource[0];
    var keyMap = {
        id: 'id', 
        name: '名字',
        age: '年龄',
      };
    dataSource.unshift(keyMap);

    dataSource.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
        v: v[k],
        position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
    }))).reduce((prev, next) => prev.concat(next)).forEach((v, i) => dataSource[v.position] = {
        v: v.v
    });

    var outputPos = Object.keys(dataSource); //设置区域,比如表格从A1到D10
    SheetNames: ['mySheet1', ...], //保存的表标题
    Sheets: {
      'mySheet1': Object.assign({},
          dataSource, //内容
          {
              '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] //设置填充区域
          })
      }
     ...
    }
    ```

## 表格对象
    每个不以!映射到单元格的键（使用A-1表示法）
    sheet[address] 返回指定地址的单元格对象。
    特殊表单键（可访问sheet[key]，每个都以!）开头：

   - sheet['!ref']：基于A-1的范围表示工作表范围。使用工作表的函数应使用此参数来确定范围。不处理在范围之外分配的单元格。

   - sheet['!margins']：表示页边距的对象。默认值遵循Excel的“正常”预设。
    Excel还具有“宽”和“窄”预设，但它们存储为原始测量值。主要属性如下：

```jsx
    / *将工作表单设置为“normal” * / 
    ws [ “！margin ” ] = {left：0.7，right：0.7，top：0.75，bottom：0.75，header：0.3，footer：0.3 }

    / *将工作表设置为“wide” * / 
    ws [ “！margin ” ] = {left：1.0，right：1.0，top：1.0，bottom：1.0，标题：0.5，页脚：0.5 }

    / *将工作表单设置为“narrow” * / 
    ws [ “！margin ” ] = {left ：0.25，right ：0.25，top ：0.75，bottom ：0.75，header ：0.3，footer ：0.3 }
 ```
 - ws['!cols']： 列属性对象的数组。
 - ws['!rows']：行属性对象的数组。
 - ws['!merges']：与工作表中合并的单元格对应的范围对象数组。 

 ## 写选项（导出write和writeFile函数接受options参数）
 - type		------- 输出数据编码（参见下面的输出类型）
 - cellDates ----	将日期存储为类型d（默认为n）
 - bookSST	---- 生成共享字符串表**
 - bookType	------	工作簿类型（默认"xlsx“）
    
## 样式

  - 单元格样式样式有fill，font，numFmt，alignment，和border。

    ```jsx
    wb["B1"].s = {
      font: { 
        sz: 14,
        bold: true,
        color: { rgb: "88FFAA99" }
        }, 
      fill: { 
        bgColor: { indexed: 64 }, 
        fgColor: { rgb: "88FF88" } 
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

## 合并单元格

  - c为列， r为行 从0开始

    ```jsx
     wb["!merges"] = [{
        s: { c: C, r: R },
        e: { c: C, r: R }
    }]
    ```

## Demo

  - 下面是个简单的demo

    ```jsx
        var tmpdata = json[0];
        json.unshift({});
        var keyMap = []; //获取keys
        for (var k in tmpdata) {
            keyMap.push(k);
            json[0][k] = k;
        }
        var tmpdata = [];//用来保存转换好的json
        json.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
            v: v[k],
            position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
        }))).reduce((prev, next) => prev.concat(next)).forEach((v, i) => tmpdata[v.position] = {
            v: v.v
        });
        var outputPos = Object.keys(tmpdata); //设置区域,比如表格从A1到D10
        tmpdata["B1"].s = { font: { sz: 14, bold: true, color: { rgb: "88FFAA99" } }, fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "88FF88" } } };//<====设置xlsx单元格样式
        tmpdata["B1"].l = { Target: "https://github.com/SheetJS/js-xlsx#writing-options", Tooltip: "Find us @ SheetJS.com!" };
        tmpdata ['C2'].l  = {Target: "＃A2" };
        tmpdata["!merges"] = [{
            s: { c: 1, r: 0 },
            e: { c: 4, r: 0 }
        }];//<====合并单元格
        var tmpWB = {
            SheetNames: ['mySheet', 'mySheet2'], //保存的表标题
            Sheets: {
                'mySheet': Object.assign({},
                    tmpdata, //内容
                    {
                        '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] //设置填充区域
                    }),
                'mySheet2': Object.assign({},
                    tmpdata, //内容
                    {
                        '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 2] //设置填充区域
                    })
            }
        };
        tmpDown = new Blob([s2ab(XLSX.write(tmpWB,
            { bookType: (type == undefined ? 'xlsx' : type), bookSST: false, type: 'binary' }//这里的数据是用来定义导出的格式类型
        ))], {
                type: ""
            });
         //创建二进制对象写入转换好的字节流
        var href = URL.createObjectURL(tmpDown); //创建对象超链接
        var downloadEle = document.createElement('a');
        downloadEle.href = href;
        downloadEle.download = "导出报表.xlsx";
        document.body.appendChild(downloadEle);
        downloadEle.click();
        window.requestAnimationFrame(function(){
            document.body.removeChild(downloadEle);
            URL.revokeObjectURL(tmpDown); //用URL.revokeObjectURL()来释放这个object URL
        });
      }
    ```
 - 处理方法.

    ```jsx
    // 字符串转字符流
    function s2ab(s) {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
    function getCharCol(n) {
      let temCol = '',
          s = '',
          m = 0·
      while (n > 0) {
          m = n % 26 + 1
          s = String.fromCharCode(m + 64) + s
          n = (n - m) / 26
      }
      return s
    }
    ```