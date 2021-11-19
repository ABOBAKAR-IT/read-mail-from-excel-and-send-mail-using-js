
var Excel = require('exceljs');
var wb = new Excel.Workbook();
var path = require('path');
var filePath = path.resolve(__dirname,'EXCEL_FILE.xlsx');
const nodemailer=require('nodemailer')
async function file_read()
{
  data= await wb.xlsx.readFile(filePath).then(async function(){

    var sh = wb.getWorksheet("Sheet1");
    sh.getRow(1).getCell(2).value;
    wb.xlsx.writeFile("sample2.xlsx");
   

    console.log(sh.rowCount);
 
   let name
   let email
   let sms
    for (let i = 2; i <= sh.rowCount; i++) {
      
         name=sh.getRow(i).getCell(2).value
         email=sh.getRow(i).getCell(3).value
         sms=sh.getRow(i).getCell(4).value
    //   console.log(name+" "+email+" "+sms);
       try {
    
        const ok=await emailsender(name,email,sms)
        if(ok)
        console.log("mail send");
    } catch (error) {
        console.log(error);
    }
    }

});
}


function emailsender(name,email,sms){

    let mailTransporter=nodemailer.createTransport({
        service:"gmail",
        host:"smtp.gmail.com",
        port:587,//465
        secure:false,
        requireTLS:true,
        auth:{
            user:"ranaabobakarit@gmail.com",
            pass:"Enter Mail Password"
        }
    });
const msg= "AOA\n rana g kya hal ha."
let mailDetails={
    from:"ranaabobakarit@gmail.com",
    to:email,
    subject:"RANA ABOBAKAR",
    text:`AOA ${name}! \n ${sms}`,
}
return mailTransporter.sendMail(mailDetails,(err,info)=>{
    if(err)
    console.log(err);
    else
    console.log("mail send");
})
}


file_read()


