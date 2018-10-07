
function saveAs(obj, fileName) {//当然可以自定义简单的下载文件实现方式 
  //创建二进制对象写入转换好的字节流
  var downloadEle = document.createElement('a');
  var objectUrl = URL.createObjectURL(obj);
  downloadEle.href = objectUrl;
  downloadEle.download = fileName;
  document.body.appendChild(downloadEle);
  downloadEle.click();
  window.requestAnimationFrame(function(){
     URL.revokeObjectURL(objectUrl); //用URL.revokeObjectURL()来释放这个object URL
    document.body.removeChild(downloadEle);
  });
  }

  var data = [{ //测试数据
      "shop": "数云食堂",
      "title": "毛衣",
      "price": "100",
      "size": "M",
  },{ //测试数据
      "shop": "数云食堂",
      "title": "毛衣",
      "price": "100",
      "size": "M",
  }];
    function downloadExl(data){
      const wopts = { bookType: 'xlsx', type: 'binary' };
      const wb = { SheetNames: ['Sheet1','Sheet2', 'Sheet3'], Sheets: {}, Props: {} };
      wb.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(data);//通过json_to_sheet转成单页(Sheet)数据
      wb.Sheets['Sheet2'] = XLSX.utils.json_to_sheet(data);
      wb.Sheets['Sheet3'] = XLSX.utils.json_to_sheet(data);
      const Blobs = new Blob([s2ab(XLSX.write(wb, wopts))], { type: "application/octet-stream" });
      saveAs(Blobs, "这里是下载的文件名" + '.' + (wopts.bookType=="biff2"?"xls":wopts.bookType));
    }

  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
