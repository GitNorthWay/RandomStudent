// Method to upload a valid excel file
function upload() {
  var files = document.getElementById('file_upload').files;
  if(files.length==0){
    alert("Please choose any file...");
    return;
  }
  var filename = files[0].name;
  var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
  if (extension == '.XLS' || extension == '.XLSX') {
      //Here calling another method to read excel file into json
      excelFileToJSON(files[0]);
  }else{
      alert("Please select a valid excel file.");
  }
}


//Method to read excel file and convert it into JSON
 function excelFileToJSON(file){
     try {
       var reader = new FileReader();
       reader.readAsBinaryString(file);
       reader.onload = function(e) {

           var data = e.target.result;
           var workbook = XLSX.read(data, {
               type : 'binary'
           });
           var result = {};
           var firstSheetName = workbook.SheetNames[0];
           //reading only first sheet data
           var jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
           //displaying the json result into HTML table
           displayJsonToHtmlTable(jsonData);
           }
       }catch(e){
           console.error(e);
       }
 }


 //Method to display the data in HTML Table
 function displayJsonToHtmlTable(jsonData){
   var table=document.getElementById("display_excel_data");
   var randomNumber = Math.floor(Math.random() * jsonData.length);
   if(jsonData.length>0){
       var htmlData='<thead><tr><th>Student</th><th>Group</th></tr></thead><tbody>';
       for(var i=0;i<jsonData.length;i++){
           var row=jsonData[i];
           if(randomNumber===i){
           htmlData+='<tr class="table-success"><td>'+'<strong>Show your knowledge: </strong>'+row["Student"]+'</td><td>'+row["Group"]
                 +'</td></tr>';
                 }
            else {
              htmlData+='<tr><td>'+row["Student"]+'</td><td>'+row["Group"]
                    +'</td></tr>';
            }
       }
       table.innerHTML=htmlData;
   }else{
       table.innerHTML='There is no data in Excel';
   }
 }
