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
## Alignment 代码对齐

  - 遵循以下的JSX语法缩进/格式. eslint: [`react/jsx-closing-bracket-location`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-closing-bracket-location.md)

    ```jsx
    // bad
    <Foo superLongParam="bar"
         anotherSuperLongParam="baz" />

    // good, 有多行属性的话, 新建一行关闭标签
    <Foo
      superLongParam="bar"
      anotherSuperLongParam="baz"
    />

    // 若能在一行中显示, 直接写成一行
    <Foo bar="bar" />

    // 子元素按照常规方式缩进
    <Foo
      superLongParam="bar"
      anotherSuperLongParam="baz"
    >
      <Quux />
    </Foo>
    ```

## Quotes 单引号还是双引号

  - 对于JSX属性值总是使用双引号(`"`), 其他均使用单引号(`'`). eslint: [`jsx-quotes`](http://eslint.org/docs/rules/jsx-quotes)

  > 为什么? HTML属性也是用双引号, 因此JSX的属性也遵循此约定.

    ```jsx
    // bad
    <Foo bar='bar' />

    // good
    <Foo bar="bar" />

    // bad
    <Foo style={{ left: "20px" }} />

    // good
    <Foo style={{ left: '20px' }} />
    ```

## Spacing 空格

  - 总是在自动关闭的标签前加一个空格，正常情况下也不需要换行. eslint: [`no-multi-spaces`](http://eslint.org/docs/rules/no-multi-spaces), [`react/jsx-tag-spacing`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-tag-spacing.md)

    ```jsx
    // bad
    <Foo/>

    // very bad
    <Foo                 />

    // bad
    <Foo
     />

    // good
    <Foo />
    ```

  - 不要在JSX `{}` 引用括号里两边加空格. eslint: [`react/jsx-curly-spacing`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-curly-spacing.md)

    ```jsx
    // bad
    <Foo bar={ baz } />

    // good
    <Foo bar={baz} />
    ```

## Props 属性

  - JSX属性名使用骆驼式风格`camelCase`.

    ```jsx
    // bad
    <Foo
      UserName="hello"
      phone_number={12345678}
    />

    // good
    <Foo
      userName="hello"
      phoneNumber={12345678}
    />
    ```

  - 如果属性值为 `true`, 可以直接省略. eslint: [`react/jsx-boolean-value`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-boolean-value.md)

    ```jsx
    // bad
    <Foo
      hidden={true}
    />

    // good
    <Foo
      hidden
    />
    ```

  - `<img>` 标签总是添加 `alt` 属性. 如果图片以presentation(感觉是以类似PPT方式显示?)方式显示，`alt` 可为空, 或者`<img>` 要包含`role="presentation"`. eslint: [`jsx-a11y/alt-text`](https://github.com/evcohen/eslint-plugin-jsx-a11y/blob/master/docs/rules/alt-text.md)

    ```jsx
    // bad
    <img src="hello.jpg" />

    // good
    <img src="hello.jpg" alt="Me waving hello" />

    // good
    <img src="hello.jpg" alt="" />

    // good
    <img src="hello.jpg" role="presentation" />
    ```

  - 不要在 `alt` 值里使用如 "image", "photo", or "picture"包括图片含义这样的词， 中文也一样. eslint: [`jsx-a11y/img-redundant-alt`](https://github.com/evcohen/eslint-plugin-jsx-a11y/blob/master/docs/rules/img-redundant-alt.md)

  > 为什么? 屏幕助读器已经把 `img` 标签标注为图片了, 所以没有必要再在 `alt` 里说明了.

    ```jsx
    // bad
    <img src="hello.jpg" alt="Picture of me waving hello" />

    // good
    <img src="hello.jpg" alt="Me waving hello" />
    ```

  - 使用有效正确的 aria `role`属性值 [ARIA roles](https://www.w3.org/TR/wai-aria/roles#role_definitions). eslint: [`jsx-a11y/aria-role`](https://github.com/evcohen/eslint-plugin-jsx-a11y/blob/master/docs/rules/aria-role.md)

    ```jsx
    // bad - not an ARIA role
    <div role="datepicker" />

    // bad - abstract ARIA role
    <div role="range" />

    // good
    <div role="button" />
    ```

  - 不要在标签上使用 `accessKey` 属性. eslint: [`jsx-a11y/no-access-key`](https://github.com/evcohen/eslint-plugin-jsx-a11y/blob/master/docs/rules/no-access-key.md)

  > 为什么? 屏幕助读器在键盘快捷键与键盘命令时造成的不统一性会导致阅读性更加复杂.

  ```jsx
  // bad
  <div accessKey="h" />

  // good
  <div />
  ```
  - 避免使用数组的index来作为属性`key`的值，推荐使用唯一ID. ([为什么?](https://medium.com/@robinpokorny/index-as-a-key-is-an-anti-pattern-e0349aece318))

  ```jsx
  // bad
  {todos.map((todo, index) =>
    <Todo
      {...todo}
      key={index}
    />
  )}

  // good
  {todos.map(todo => (
    <Todo
      {...todo}
      key={todo.id}
    />
  ))}
  ```

  - 对于所有非必须的属性，总是手动去定义`defaultProps`属性.

  > 为什么? propTypes 可以作为模块的文档说明, 并且声明 defaultProps 的话意味着阅读代码的人不需要去假设一些默认值。更重要的是, 显示的声明默认属性可以让你的模块跳过属性类型的检查.

  ```jsx
  // bad
  function SFC({ foo, bar, children }) {
    return <div>{foo}{bar}{children}</div>;
  }
  SFC.propTypes = {
    foo: PropTypes.number.isRequired,
    bar: PropTypes.string,
    children: PropTypes.node,
  };

  // good
  function SFC({ foo, bar, children }) {
    return <div>{foo}{bar}{children}</div>;
  }
  SFC.propTypes = {
    foo: PropTypes.number.isRequired,
    bar: PropTypes.string,
    children: PropTypes.node,
  };
  SFC.defaultProps = {
    bar: '',
    children: null,
  };
  ```

## Refs

  - 总是在Refs里使用回调函数. eslint: [`react/no-string-refs`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/no-string-refs.md)

    ```jsx
    // bad
    <Foo
      ref="myRef"
    />

    // good
    <Foo
      ref={(ref) => { this.myRef = ref; }}
    />
    ```


## Parentheses 括号

  - 将多行的JSX标签写在 `()`里. eslint: [`react/jsx-wrap-multilines`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-wrap-multilines.md)

    ```jsx
    // bad
    render() {
      return <MyComponent className="long body" foo="bar">
               <MyChild />
             </MyComponent>;
    }

    // good
    render() {
      return (
        <MyComponent className="long body" foo="bar">
          <MyChild />
        </MyComponent>
      );
    }

    // good, 单行可以不需要
    render() {
      const body = <div>hello</div>;
      return <MyComponent>{body}</MyComponent>;
    }
    ```

## Tags 标签

  - 对于没有子元素的标签来说总是自己关闭标签. eslint: [`react/self-closing-comp`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/self-closing-comp.md)

    ```jsx
    // bad
    <Foo className="stuff"></Foo>

    // good
    <Foo className="stuff" />
    ```

  - 如果模块有多行的属性， 关闭标签时新建一行. eslint: [`react/jsx-closing-bracket-location`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-closing-bracket-location.md)

    ```jsx
    // bad
    <Foo
      bar="bar"
      baz="baz" />

    // good
    <Foo
      bar="bar"
      baz="baz"
    />
    ```

## Methods 函数

  - 使用箭头函数来获取本地变量.

    ```jsx
    function ItemList(props) {
      return (
        <ul>
          {props.items.map((item, index) => (
            <Item
              key={item.key}
              onClick={() => doSomethingWith(item.name, index)}
            />
          ))}
        </ul>
      );
    }
    ```

  - 当在 `render()` 里使用事件处理方法时，提前在构造函数里把 `this` 绑定上去. eslint: [`react/jsx-no-bind`](https://github.com/yannickcr/eslint-plugin-react/blob/master/docs/rules/jsx-no-bind.md)

  > 为什么? 在每次 `render` 过程中， 再调用 `bind` 都会新建一个新的函数，浪费资源.

    ```jsx
    // bad
    class extends React.Component {
      onClickDiv() {
        // do stuff
      }

      render() {
        return <div onClick={this.onClickDiv.bind(this)} />;
      }
    }

    // good
    class extends React.Component {
      constructor(props) {
        super(props);

        this.onClickDiv = this.onClickDiv.bind(this);
      }

      onClickDiv() {
        // do stuff
      }

      render() {
        return <div onClick={this.onClickDiv} />;
      }
    }
    ```·
