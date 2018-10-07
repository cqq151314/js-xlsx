function saveAs(obj, fileName) {
  var tmpa = document.createElement("a");
  tmpa.download = fileName || "下载";
  tmpa.href = URL.createObjectURL(obj);
  tmpa.click();
  setTimeout(function () {
      URL.revokeObjectURL(obj);
  }, 100);
}
var data = [{ //测试数据
    "shop": "数云食堂",
    "title": "毛衣",
    "price": "100",
    "size": "M",
  },{ //测试数据
    "shop": "数云食堂",
    "title": "毛衣毛衣",
    "price": "100",
    "size": "M",
  }];
  
function downloadExl(json, type) {
  var tmpdata = [];//用来保存转换好的json
  var keyMap = ["shop", "title", "price", "size"];
  
  json.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
      v: v[k],
      position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
  }))).reduce((prev, next) => prev.concat(next)).forEach((v, i) => tmpdata[v.position] = {
      v: v.v
  });

  var outputPos = Object.keys(tmpdata); //设置区域,比如表格从A1到D10
  var wopts = { bookType: 'xlsx', type: 'binary' };
  var wbs = { SheetNames: ['Sheet1'], Sheets: {
    "Sheet1": Object.assign({},
      tmpdata, //内容
      {
          '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] //设置填充区域
      })
    }
  };
  // 设置单元格样式
  wbs.Sheets['Sheet1']["B2"].s = {
    font: {
        sz: 20,
        bold: true,
        color: { rgb: "88FFAA99" },
        name: '宋体',
        italic: false,
        underline: true
      },
    fill: {
        fgColor: { rgb: "#000000" },
    },
    alignment: {
        horizontal: "center" ,
        vertical: "center",
        wrapText: true
    },
    border:{
        top:{
            style:'thick',
            color: { auto: 2}
        },
        left: {
          style:'dotted',
          color: { rgb: "88FFAA99" }
        },
        diagonal:{
            style:'thick',
            color: { rgb: "88FFAA99" }
        },
        bottom:{
            style:'thick',
            color: { theme: "1", tint: "-0.9"},
        },
        right:{
            style:'thick',
            color: { indexed: 64}
        },
        diagonalUp:	true, // 对角线
        diagonalDown: true
    }
};

  var Blobs = new Blob([s2ab(XLSX.write(wbs, wopts))], { type: "application/octet-stream" });
  saveAs(Blobs, "这里是下载的文件名" + '.' + (wopts.bookType=="biff2" ? "xls" : wopts.bookType));
};

// 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
function getCharCol(n) {
  let temCol = '',
      s = '',
      m = 0
  while (n > 0) {
      m = n % 26 + 1
      s = String.fromCharCode(m + 64) + s
      n = (n - m) / 26
  }
  return s
}

 //字符串转字符流
function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}