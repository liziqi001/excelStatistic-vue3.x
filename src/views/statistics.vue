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
        <h3>统计属性：</h3>
        属性名称：
        <Input v-model="addProp" placeholder="" style="width: 200px" />
        统计的原表属性：
        <Select v-model="addByProp" style="width:200px">
            <Option v-for="item in tableColumns" :value="item.key" :key="item.key">{{ item.key }}</Option>
        </Select>
        计算方式：
        <Select v-model="addType" style="width:200px">
            <Option  value="plus" >和</Option>
            <Option  value="percent" >百分比</Option>
            <Option  value="average" >平均值</Option>
        </Select>
        <div v-show="addType=='percent'&&addByProp!=''">
          条件值：
          <Select v-model="addValue" style="width:200px">
              <Option v-for="item in columnsValue" :value="item" :key="item">{{ item }}</Option>
          </Select>
        </div>
        <Button @click="add">确认</Button>
    </div>
    <br><hr>
 基于导入表的第一列分类：
    <div id="table2" >
      <Table
        height="300"
        :columns="newTableColumns"
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
      newTableColumns:[],

      typeObj:{},//按对象第一列分类{'同济'：[{},{}],'协和'：[{},{}]}
      addProp:'',
      addByProp:'',
      addType:'',
      addValue:'',
    };
  },
  computed: {
    firstKey(){
      return Object.keys(this.tableData[0])[0]
    },
    columnsValue(){//单列的值去重[]
      let arr=[] 
      for(let i in this.tableData){
        arr.push(this.tableData[i][this.addByProp])
      }
      return [...new Set(arr)]
    },
    
  },
  methods: {
    add(){//添加导出表的属性
      let addProp=this.addProp;
      let addByProp=this.addByProp;
      let addType=this.addType;
      let addValue=this.addValue;
      if(this.newTableData.length>0){
        let hasProp=Object.keys(this.newTableData[0])
        if(hasProp.indexOf(addProp)>-1){
          this.$Message.warning('属性名称不能重复！')
          return;
        }
      }
      if(addProp==''||addByProp==''||addType==''||(addType=='percent'&&addProp=='')){
        this.$Message.warning('请填写完整！')
        return;
      }else{
        this.newTableColumns.push({title:addProp,key:addProp})
        var length=Object.keys(this.typeObj).length
        if(addType=='plus'){//和  
          for(let i in this.typeObj){
            var plus=0;
            var objs=this.typeObj[i]
            for(let o of objs){
              plus+=parseInt(o[addByProp])
            }

            if(this.newTableData.length<length){
              var obj={};
              obj[this.firstKey]=i;
              obj[addProp]=plus;
              this.newTableData.push(obj);
            }else{
              for(let o of this.newTableData){
                if(i==o[this.firstKey]){
                  o[addProp]=plus;
                }
              }
            }
          }
        }else if(addType=='percent'){
          for(let i in this.typeObj){
            var valueCount=0;
            var objs=this.typeObj[i]
            for(let o of objs){
              if(o[addByProp]==addValue){
                valueCount++;
              }
            }
            var percent=(valueCount/objs.length).toFixed(4)*100+'%';
            if(this.newTableData.length<length){
              var obj={};
              obj[this.firstKey]=i;
              obj[addProp]=percent;
              this.newTableData.push(obj);
            }else{
              for(let o of this.newTableData){
                if(i==o[this.firstKey]){
                  o[addProp]=percent;
                }
              }
            }
          }
        }else if(addType=='average'){
          for(let i in this.typeObj){
            var plus=0;
            var objs=this.typeObj[i]
            for(let o of objs){
              plus+=parseInt(o[addByProp])
            }
            var average=(plus/objs.length).toFixed(2);
            if(this.newTableData.length<length){
              var obj={};
              obj[this.firstKey]=i;
              obj[addProp]=average;
              this.newTableData.push(obj);
            }else{
              for(let o of this.newTableData){
                if(i==o[this.firstKey]){
                  o[addProp]=average;
                }
              }
            }
          }
        }

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
            vueModel.newTableColumns.push(columns[0]);
            vueModel.tableData=json;
            vueModel.addMarkersFinished = true;

            var obj={};
            let firstValue=json[0][vueModel.firstKey]
            obj[firstValue]=[json[0]];
            for(let i in json){
              var thisObj=json[i];
              if(i!='0'){
                var contain=false;
                for(let j in obj){
                  if(thisObj[vueModel.firstKey]==j){
                    obj[j].push(thisObj)
                    contain=true;
                    break;
                  }
                }
                if(!contain){
                  obj[thisObj[vueModel.firstKey]]=[thisObj]
                }
              }
            }
            console.log(obj)
            vueModel.typeObj=obj;
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
