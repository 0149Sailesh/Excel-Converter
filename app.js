import express from "express";
import path from "path";
import { fileURLToPath } from 'url';
import fileUpload from "express-fileupload";
import csvtojson from "csvtojson"
import csvToJson from "convert-csv-to-json"
import * as XLSX from 'xlsx';
import * as fs from "fs";
XLSX.set_fs(fs);

const __filename = fileURLToPath(import.meta.url);

const __dirname = path.dirname(__filename);


console.log(__dirname)

const app = express();

app.use(express.json())

app.use(fileUpload())

app.set('view engine', 'ejs')

app.use(express.static(__dirname + '/public'));

app.use("/css",
    express.static(path.join(__dirname, "node_modules/bootstrap/dist/css"))
  )
app.use("/js",
    express.static(path.join(__dirname, "node_modules/bootstrap/dist/js"))
  )
app.use("/js", express.static(path.join(__dirname, "node_modules/jquery/dist")))

const convertToTime = (sum)=>{

   return Math.floor(String(sum/60))+":"+ String(("0"+(sum%60)).slice(-2));

}

const checkVertical = (workbook, x)=>{
    let flag = false;

    let i;

    for(i=0;i<30;i++){ 
        let num = Number(x.substring(1)) + i;
        let cell = x[0]+num;

        if(!workbook[cell]){
            continue;
        } else {
            flag = true;
            return flag;
        }
    }

    return flag;

    //flag = true means not the end
    //flag = false means the end
}

const checkHorizontal = (workbook, x)=>{
    let flag = false;
    let i,k;
    for(i = x.charCodeAt(1), k=0; k<30 ; i++,k++){

        if(i>90){
            i = 64;
            x = (String.fromCharCode(x.charCodeAt(0)+1)) + x.slice(1);
            continue;
        }
        let cell;
        if(x[0]=='@'){
            cell = String.fromCharCode(i) + x.substring(2);
        } else {
            cell = x[0] + String.fromCharCode(i) + x.substring(2);
        }

        if(!workbook[cell]){
            continue;
        } else {
            flag = true;
            return true;
        }
    }

    return false;

     //flag = true means not the end
    //flag = false means the end
}
const calcHelper = (workbook, x) =>{
    let i;
   let sum = 0;
   let count = 0;
   x = '@'+x; //@ has ASCI code 64

   for(i = x.charCodeAt(1); ;i++){
    
        if(i>90){
            i = 64;
            x = (String.fromCharCode(x.charCodeAt(0)+1)) + x.slice(1);
            continue;
        }
    let cell;
    if(x[0]=='@'){
        cell = String.fromCharCode(i) + x.substring(2);
    } else {
        cell = x[0] + String.fromCharCode(i) + x.substring(2);
    }
    // console.log("x val: ",x)
    // console.log("Cell accessed: ", cell)
    if(!workbook[cell]){
        //Don't forget both continue and break
        let checkEnd = checkHorizontal(workbook, x);

        if(checkEnd) continue;
        else break;
    }
    const time = workbook[cell].v.split(":") ;

    let hours = Number(time[0])*60;
    let minutes = Number(time[1]);

    sum+= (hours+minutes);
    count++;
   }

   let avg = Math.floor(sum/count);
   
   return [sum,avg];
}

const eachPerson = (workbook, x)=>{
    var nameOfPerson;
    //Get name of Person
    if(x=='B7'){
        nameOfPerson = workbook['A5'].v;
    } else {
        nameOfPerson = workbook['A'+ (Number(x.substring(1))-1)].v;
    }

    let nameExtracted = nameOfPerson.split("Name : ").pop().split(" CardNo").shift();
    let cardNo = nameOfPerson.split("CardNo : ").pop().split("Present").shift()

    console.log(nameExtracted);
    console.log(cardNo)
   
    //Fetch in time parameter [sum, avg]
    let inTime = calcHelper(workbook, x);
    let inTimeAvg = convertToTime(Number(inTime[1]));
    console.log("In time:", inTimeAvg);

    //Fetch out time parameters [sum, avg]
    let outTime = calcHelper(workbook, x[0]+(Number(x.substring(1))+1));
    let outTimeAvg = convertToTime(Number(outTime[1]));
    console.log("Out time:", outTimeAvg);

    //Calculate sum of all working days
    let sumOfAllDays = (Number(outTime[0]) - Number(inTime[0]));
    let sumInTime = convertToTime(sumOfAllDays);
    console.log("Sum of working hours: ", sumInTime);

    return [nameExtracted, inTimeAvg, outTimeAvg, sumInTime, cardNo]

}

const itThruList = (workbook) =>{
    let firstCell = 'B7';
    let jsonData = [];
    
    let i;
    for(i=0;;i++){
        let curCellNo = Number(firstCell[1]) + i*8;
        let curCell = 'B'+curCellNo;

        //Write code here to check for vertical end
        if(!workbook[curCell]){

            let checkVerticalEnd = checkVertical(workbook, curCell);

            if(checkVerticalEnd) continue;
            else break;
        }

        let a = eachPerson(workbook, curCell);

        jsonData.push({
            "Name of Person": a[0],
            "Card No": a[4],
            "Avg In Time": a[1],
            "Avg out time": a[2],
            "Present/Absent": a[3]
        })
    }
    return jsonData
}

app.get('/', (req, res)=>{
    //var workbook = XLSX.readFile("PerformanceRegister-FACULTY-EEE.xlsx")
    res.render('index')
})

app.post('/',(req, res)=>{

    var workbook = XLSX.read(req.files.userFile.data);

    let jsonData = itThruList(workbook.Sheets.Sheet1);
    
    console.log(jsonData)

    const newWb = XLSX.utils.book_new();
    const fileName = "sample"
    const datasheet = XLSX.utils.json_to_sheet(jsonData)

    XLSX.utils.book_append_sheet(newWb, datasheet, fileName.replace("/", ""))

    const binaryWorkbook = XLSX.write(newWb, {
        type: "buffer",
        bookType: "xlsx",
      });

      res.setHeader(
        "Content-Disposition",
        'attachment; filename="SheetJSNode.xlsx"'
      );
    
      res.setHeader("Content-Type", "application/vnd.ms-excel");
      
    return res.status(200).send(binaryWorkbook);
   


})
app.listen(3000, ()=>{
    console.log("Example app is running on port 3000")
})