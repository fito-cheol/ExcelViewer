<template>
  <div align="left" >
    <input type="file" id="excelFile" @change="excelExport"/>
    <br/>
    <span> {{header}}</span> <br/>
    <div v-for="(row,index) in rows" :key="index" >
      <span> {{row}}</span> <br/>
    </div>
    <button @click="clear"> Clear </button>
  </div>

</template>

<script>
const XLSX = require('xlsx');
export default {
  name: 'HelloWorld',
  data(){
    return{
      header:"",
      rows:[]
    }
  },
  methods:{
    clear(){
      this.header = "";
      this.rows = [];
    },
    excelExport(event){
      let input = event.target;
      let reader = new FileReader();
      reader.onload = function(){
        let fileData = reader.result;
        let wb = XLSX.read(fileData, {type : 'binary'});
        wb.SheetNames.forEach(function(sheetName){
          const sheet = wb.Sheets[sheetName]
          const rowObj =XLSX.utils.sheet_to_json(sheet);
          const header = this.get_header_row(sheet)

          console.log('Header:', header)
          console.log('Rows:', rowObj)
          this.header = header
          this.rows = rowObj
        }.bind(this))
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
