<style scoped>
#tool {
  position: absolute;
  top: 60px;
  left: 0;
  bottom: 0;
  width: 100%;
  padding:20px;
  background: rgb(190, 191, 211);
}
#upload {
  width: 90%;
  height: 50px;
}
#table {
  width: 90%;
  height: 300px;
  background: cornflowerblue;
}
#add {
    margin:10px;
  width: 90%;
  height: 50px;
  background: cornflowerblue;
}
#table2 {
    margin:10px;
  width: 90%;
  height: 300px;
  background: cornflowerblue;
}
</style>
<template>
  <div id="tool">
    <div id="upload">
      <input
        type="file"
        @change="importFile(this)"
        id="imFile"
        style="display: none"
        accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
      />
      <Button type="primary" :loading="parseExcel_loading" style="float:left" @click="uploadFile()">
        <span v-if="!parseExcel_loading">导 入</span>
        <span v-else>解析中...</span>
      </Button>
      <div style="float:left">
        <Alert type="error" show-icon v-show="uploadFailed">上传失败：{{uploadMsg}}</Alert>
        <Alert type="success" show-icon v-show="addMarkersFinished">解析成功：共{{tableData.length}}行数据</Alert>
      </div>
    </div>
    <div id="table">
        <Table
            height="300"
            :columns="tableColumns"
            :data="tableData"
            size="small"
            ref="table"
            @on-selection-change="tableSelected"
        ></Table>
    </div>
    <div class='add'>
        <h3>选择拆分列：</h3>
        列名：
        <Select v-model="type" multiple @input="limit" style="width:200px">
            <Option  v-for="v in tableColumns"  :value="v.title" :key="v.title">{{v.title}}</Option>
        </Select>
        分割符：
        <Input v-model="joinStr" placeholder="" style="width: 200px" />
        <Button @click="split">确认</Button>
    </div>
    <br><hr>
    结果：
    <div id="table2" >
      <Table
        height="300"
        :columns="tableColumns"
        :data="newTableData"
        size="small"
        ref="table2"
        @on-selection-change="tableSelected"
      ></Table>
      <Button type="primary" size="large" @click="exportData(1)">
        <Icon type="ios-download-outline"></Icon>导出数据
      </Button>
    </div>
  </div>
</template>
<script type="text/ecmascript-6">
import XLSX from "xlsx";
import $ from "jquery";

export default {
  data() {
    return {
      parseExcel_loading: false, // 解析excel中。。。
      imFile: "", // 导入文件el
      uploadFailed: false,
      uploadMsg: "",
      addMarkersFinished: false,

      tableData: [],

      newTableData:[],
      selection: [],

      tableColumns:[],
      //newTableColumns:[],

      joinStr:'',
      type:[],
    };
  },
  computed: {
 
  },
  methods: {
    limit(e) {
      if (e.length > 2) {
        this.$Message.warning('最多选择2列');
        e.pop();
      }
    },
    split(){//
      let joinStr=this.joinStr;
      let type=this.type;
      if(joinStr==''||type.length==0){
          this.$Message.warning('请填写完整！')
          return;
      }

      var firstSplitArr=[];
      var newObj=[];
      for(let index in type){
        var d=index==0?this.tableData:firstSplitArr
        for(let t of d){
          var splitArr=t[type[index]].split(joinStr);
          for(let i of splitArr){
            var obj={...t};
            obj[type[index]]=i;
            if(index==0){
              firstSplitArr.push(obj)
            }else{
              newObj.push(obj)
            }
          }
        } 
      }
      if(type.length==1){
        this.newTableData=firstSplitArr
      }else{
        this.newTableData=newObj;
      }
      
    },
    uploadFile: function() {
      // 点击导入按钮
      this.imFile.click();
    },
    importFile: function() {
      // 导入excel
      let obj = this.imFile;
      if (!obj.files) {
        this.uploadFailed = true;
        $("#imFile")[0].value("");
        return;
      }
      var f = obj.files[0];
      if (
        f &&
        f.name.indexOf(".xlsx") < 0 &&
        f.name.indexOf(".xls") < 0 &&
        f.name.indexOf(".csv") < 0
      ) {
        this.uploadFailed = true;
        this.uploadMsg = "文件格式不正确";
        $("#imFile")[0].value("");
        return;
      }
      //初始化
      this.parseExcel_loading = true;
      this.addMarkersFinished = false;
      this.uploadFailed = false;
      this.tableData = [];
      this.lnglats = [];

      var reader = new FileReader();
      let vueModel = this;
      reader.onload = function(e) {
        var data = e.target.result;
        if (vueModel.rABS) {
          vueModel.wb = XLSX.read(btoa(this.fixdata(data)), {
            // 手动转化
            type: "base64"
          });
        } else {
          vueModel.wb = XLSX.read(data, {
            type: "binary"
          });
        }
        let json = XLSX.utils.sheet_to_json(
          vueModel.wb.Sheets[vueModel.wb.SheetNames[0]]
        );
        $("#imFile").val("");
        vueModel.parseExcel_loading=false;
        console.log(json);
        if (vueModel.analyzeData(json)) {//有数据
            let columns=[];
            let data=[];
            for(let m in json[0]){
                columns.push({key:m,title:m})
            }
            vueModel.tableColumns=columns;
            vueModel.tableData=json;
            vueModel.addMarkersFinished = true;

        } else {
        }
      };
      if (this.rABS) {
        reader.readAsArrayBuffer(f);
      } else {
        reader.readAsBinaryString(f);
      }
    },
    analyzeData(json) {
      if (json.length != 0) {
        let equal=true;
        for (let i in json) {
          if (parseInt(i) + 1 < json.length) {
            if (Object.keys(json[i]).length != Object.keys(json[parseInt(i) + 1]).length) {
                this.uploadFailed = true;
                this.uploadMsg = "数据缺失";
                equal=false;
            }
          }
        }
        return equal;
      } else {
        this.uploadFailed = true;
        this.uploadMsg = "excel无数据";
         return false
        //$('#imFile').val('');
      }
    },

    s2ab: function(s) {
      // 字符串转字符流
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i = 0; i !== s.length; ++i) {
        view[i] = s.charCodeAt(i) & 0xff;
      }
      return buf;
    },
    getCharCol: function(n) {
      // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
      let s = "";
      let m = 0;
      while (n > 0) {
        m = (n % 26) + 1;
        s = String.fromCharCode(m + 64) + s;
        n = (n - m) / 26;
      }
      return s;
    },
    fixdata: function(data) {
      // 文件流转BinaryString
      var o = "";
      var l = 0;
      var w = 10240;
      for (; l < data.byteLength / w; ++l) {
        o += String.fromCharCode.apply(
          null,
          new Uint8Array(data.slice(l * w, l * w + w))
        );
      }
      o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
      return o;
    },
    exportData(type) {
      //iview导出
      if (type === 1) {
        this.$refs.table2.exportCsv({
          filename: "The original data",
          data:this.newTableData.filter((data,index)=>{
            for(let i in data){
              data[i]= "\t" +data[i].toString()
            }
            return data;
          }),
        });
      }
    },
    tableSelected(selection) {
      this.selection = selection;
    }
  },
  mounted() {
    this.imFile = document.getElementById("imFile");
    this.outFile = document.getElementById("downlink");
  }
};
</script>
