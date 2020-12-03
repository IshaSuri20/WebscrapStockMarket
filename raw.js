let fs = require("fs");
let path = require("path");
let req = require("request");
let cheerio = require("cheerio");
let xlsx = require("xlsx");
req("https://www.moneycontrol.com/markets/indian-indices/",find);
function find(err,res,html){
    console.log(res.statusCode);
 let fullhtml="<table>";
    let ch = cheerio.load(html);
    let detailsElement = ch("div#indices_stocks.right_block");
 for (let j = 0; j < detailsElement.length; j++) {
     let heading=ch(detailsElement[j]).find("h1.blue_txt").text();
     let details=ch(detailsElement[j]).find("table.responsive").find("tbody tr");
     for (let i= 0; i < details.length; i++) {
         let rowcol=ch(details[i]).find("td")
         let sname = ch(rowcol[0]).text().trim();
                let LTP = ch(rowcol[1]).text().trim();
                let change  = ch(rowcol[2]).text().trim();
                let Volume = ch(rowcol[3]).text().trim();
                let BuyPrice = ch(rowcol[4]).text().trim();
                let SellPrice  = ch(rowcol[5]).text().trim();
                let buyQTY = ch(rowcol[6]).text().trim();
                let SellQty = ch(rowcol[7]).text().trim();
               // console.log(`name:${sname}\nLTP:${LTP}\nchange:${change}\nVolume:${Volume}\nBuyPrice:${BuyPrice}\nSellPrice:${SellPrice}\nbuyQTY:${buyQTY}\nSellQty:${SellQty}`)	
               // console.log("###############################");
                stock(heading,sname,LTP,change,Volume,BuyPrice,SellPrice,buyQTY,SellQty);
               
             }
            }  
            }
            function stock(heading,sname,LTP,change,Volume,BuyPrice,SellPrice,buyQTY,SellQty){
                let dirpath=heading;
                let stock=
            {
                NAME:sname,
                LTP:LTP,
                CHANGE:change,
                VOL:Volume,
                BUYPRICE:BuyPrice,
                SPRICE:SellPrice,
                SQTY:SellQty,
                BUYQTY:buyQTY
            }
            if(fs.existsSync(dirpath)){
             } else{
                 fs.mkdirSync(dirpath);
             }
             let stockFilePath=path.join(dirpath,sname+".xlsx");
             let pData=[];
             if(fs.existsSync(stockFilePath)){
                 pData=excelReader(stockFilePath,sname)
                 pData.push(stock);
            
             }else{
                 console.log("File of Stock",stockFilePath,"created");
                 pData=[stock];
               }
               excelWriter(stockFilePath,pData,sname);
            }
            function excelReader(filePath, name) {
               if (!fs.existsSync(filePath)) {
                   return null;
               }
              
               let wt = xlsx.readFile(filePath);
             
               let excelData = wt.Sheets[name];
              
               let ans = xlsx.utils.sheet_to_json(excelData);
              
               return ans;
            }
            function excelWriter(filePath, json, name) {
               // console.log(xlsx.readFile(filePath));
               var newWB = xlsx.utils.book_new();
               // console.log(json);
               var newWS = xlsx.utils.json_to_sheet(json)
               xlsx.utils.book_append_sheet(newWB, newWS, name)//workbook name as param
               xlsx.writeFile(newWB, filePath);
            }
                            