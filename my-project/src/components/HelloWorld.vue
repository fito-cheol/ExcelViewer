<template>
  <div align="left" >
    <form>
      <input type='reset' @click="clear"/>
      <br/>
      <input type="file" id="excelFile" @change="excelExport"/>
    </form>
    <br/>
    <div v-for="(sheetName, index) in tabs" :key="'sheet'+index">
      <h2> {{sheetName}}</h2>
      <v-data-table :items="rows[sheetName]" :headers="headers[sheetName]" />
      <hr>
    </div>
  </div>
</template>

<script>
const XLSX = require('xlsx');
export default {
  name: 'HelloWorld',
  data(){
    return{
      tabs:[],
      headers: {},
      rows: {}
    }
  },
  methods:{
    clear(){
      this.tabs=[]
      this.headers = {};
      this.rows = {};
    },
    excelExport(event){
      let input = event.target;
      let reader = new FileReader();
      reader.onload = function(){
        this.clear()
        let fileData = reader.result;
        let wb = XLSX.read(fileData, {type : 'binary'});
        this.tabs = wb.SheetNames
        wb.SheetNames.forEach(function(sheetName){
          const sheet = wb.Sheets[sheetName]
          const rowObj =XLSX.utils.sheet_to_json(sheet);
          const headers = this.get_header_row(sheet)

          // console.log('Header:', headers)
          // console.log('Rows:', rowObj)
          this.headers[sheetName] = headers.map(header => {return {value:header, text: header}})
          this.rows[sheetName] = rowObj
        }.bind(this))
        console.log('Tabs',this.tabs)
        console.log('Headers',this.headers)
        console.log('Rows', this.rows)
      }.bind(this);
      reader.readAsBinaryString(input.files[0]);
    },
    get_header_row(sheet) {
      var headers = [];
      var range = XLSX.utils.decode_range(sheet['!ref']);
      var C, R = range.s.r; /* start in the first row */
      /* walk every column in the range */
      for(C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[XLSX.utils.encode_cell({c:C, r:R})] /* find the cell in the first row */

        var hdr = "UNKNOWN " + C; // <-- replace with your desired default
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);

        headers.push(hdr);
      }
      return headers;
    }
  }


}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
h3 {
  margin: 40px 0 0;
}
ul {
  list-style-type: none;
  padding: 0;
}
li {
  display: inline-block;
  margin: 0 10px;
}
a {
  color: #42b983;
}
</style>
