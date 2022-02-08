const express = require("express");
const multer = require("multer");
const fs = require('fs');
const app = express();

// Store the file in uploads folder
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads');
    },
    filename: (req, file, cb) => {
        cb(null, "names.xlsx");
    }
})

const upload = multer({ storage: storage});
app.use(express.static('public'));
app.listen(3000, () => console.log('App is listening...'));

app.post("/upload", upload.single("excel"), function (req, res) {
    return res.redirect("http://localhost:3000/");
});

if(fs.existsSync('uploads/names.xlsx')){
  var XLSX = require("xlsx");
  var workbook = XLSX.readFile("uploads/names.xlsx");
  var sheet_name_list = workbook.SheetNames;
  console.log(sheet_name_list); // getting as Sheet1
  var headers = {};
  var data = [];
  sheet_name_list.forEach(function (y) {
    var worksheet = workbook.Sheets[y];
    //getting the complete sheet
    // console.log(worksheet);

    for (z in worksheet) {
      if (z[0] === "!") continue;
      //parse out the column, row, and value
      var col = z.substring(0, 1);
      // console.log(col);

      var row = parseInt(z.substring(1));
      // console.log(row);

      var value = worksheet[z].v;
      // console.log(value);

      //store header names
      if (row == 1) {
        headers[col] = value;
        // storing the header names
        continue;
      }

      if (!data[row]) data[row] = {};
      data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift();
    data.shift();
    console.log(data);
  });
}

app.get('/excel-api',function (req, res){
    res.json(data);
});