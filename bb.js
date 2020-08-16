const fs = require("fs");
const request = require("request-promise-native");

const prompt =  require('prompt-sync')();

var XLSX =  require('xlsx');


var cases = new Array();

async function test(){





var table = await XLSX.readFile('Links.xlsx');


var file_names = ['Links.xlsx']

var table_names = ['Sheet1']

var start_range = [10]

var stop_range = [20]

var runningtime = new Array();
var producedby = new Array();
var writtenby = new Array();
var screenplayby = new Array()
var storyby = new Array();
var starring = new Array()
var musicby = new Array();
var cinematography = new Array();
var editedby = new Array();
var distributedby = new Array();
var releasedate = new Array()
var boxoffice = new Array()




var file_names = 'Login.xlsx'



async function Report(columnValue,file,start_range,stop_range,table_name){


var sheet =await table.Sheets[table_name];

for (var i = start_range; i< stop_range;i++)
{

   	
   var cell_address = await {c:columnValue, r:i};

    var cell_ref = await XLSX.utils.encode_cell(cell_address);

    var cell = await sheet[cell_ref];

    await cases.push(cell.v);

       
    
}
    await console.log(cases);
   

}


await Report(0,file_names[0],start_range[0],stop_range[0],table_names[0]);





const {Builder, By, Key} = require('selenium-webdriver');
async function scrap(){


for (i =0; i<cases.length; i++){

const driver = new Builder().forBrowser('chrome').build();
await driver.get(cases[i]); 
await driver.manage().setTimeouts( { implicit: 20000 } );


try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//td[contains(.,'minutes')]")).getText(); runningtime.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');runningtime.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Produced by')]")).getText();producedby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');producedby.push(0)}


try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Written by')]")).getText();writtenby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');writtenby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Screenplay by')]")).getText();screenplayby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');screenplayby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Story by')]")).getText();storyby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');storyby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Starring')]")).getText();starring.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');starring.push(0)}



try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Music by')]")).getText();musicby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');musicby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Cinematography')]")).getText();cinematography.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');cinematography.push(0)}


try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Edited by')]")).getText();editedby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');editedby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Distributed by')]")).getText();distributedby.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');distributedby.push(0)}

try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Release date')]")).getText();releasedate.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');releasedate.push(0)}


try{
var element = await driver.findElement(By.xpath(".//*[@class='infobox vevent']//tr[contains(.,'Box office')]")).getText();boxoffice.push(element);
}
catch(NoSuchElementError){console.log('#No Element by Xpath[Text]');boxoffice.push(0)}





//await console.log(runningtime)
await driver.quit();

}
//await driver.quit();

file = 'Links.xlsx'
 var target_table = await XLSX.readFile(file,{cellDates: true });
    var target_sheet =await target_table.Sheets['Sheet2'];

    
    



    await XLSX.utils.sheet_add_aoa(target_sheet, [runningtime], {origin: 'B1'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [producedby], {origin: 'B2'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [writtenby], {origin: 'B3'});
    
    await XLSX.utils.sheet_add_aoa(target_sheet, [screenplayby], {origin: 'B4'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [storyby], {origin: 'B5'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [starring], {origin: 'B6'});
    
    await XLSX.utils.sheet_add_aoa(target_sheet, [musicby], {origin: 'B7'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [cinematography], {origin: 'B8'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [editedby], {origin: 'B9'});
    
    await XLSX.utils.sheet_add_aoa(target_sheet, [distributedby], {origin: 'B10'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [releasedate], {origin: 'B11'});
    await XLSX.utils.sheet_add_aoa(target_sheet, [boxoffice], {origin: 'B12'});
    
    
    
    
    
    
    
    
    await XLSX.writeFile(target_table, file);






}
await scrap();








}
test();
