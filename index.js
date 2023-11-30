const {Builder, By, Key, until,Select,WebElement,Actions} = require('selenium-webdriver');
const firefox = require('selenium-webdriver/firefox');
const ExcelJS = require('exceljs');
const options= new firefox.Options()
const downloadFilepath = "D:\\ProgData\\RPA\\reports\\"
const ipadd = "http://10.163.132.13/"
// options.addArguments('-headless')
options.addArguments('-private')
options.setPreference("browser.download.folderList",2)
options.setPreference("browser.download.manager.showWhenStarting",false)
options.setPreference("browser.download.dir",downloadFilepath)
options.setPreference("browser.helperApps.neverAsk.saveToDisk",'application/xlsx')
let driver =  new Builder().forBrowser('firefox').setFirefoxOptions(options).build();
const fs = require('fs');
const path = require('path');
let Client = require('ssh2-sftp-client');
let tries=3
let correct=false

async function navLogin(){
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[3]')).sendKeys('rpafogo');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[4]')).sendKeys('mmh');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[5]')).sendKeys('Hotel321#');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/button')).click();  
  await driver.wait(until.elementLocated(By.css('#sideMenu')),10000)
  const frameSideMenu = await driver.findElement(By.css('#sideMenu'))  
  await driver.switchTo().frame(frameSideMenu)
  await driver.wait(until.elementLocated(By.xpath('/html/body/div[3]/div[2]/a[4]')),10000).then(element=>{element.click()})
  // await driver.findElement(By.xpath('/html/body/div[3]/div[2]/a[4]')).click()  
  await new Promise(resolve => setTimeout(resolve, 10000));
  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
}
async function navReportMixItems(){
  //reporte 1 
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[7]/td[2]/h2')).click()
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[8]/td/div/table/tbody/tr[6]/td[2]/a')).click()
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[7]/td[2]/h2'))),10000).then(element=>{element.click()})
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[8]/td/div/table/tbody/tr[6]/td[2]/a'))),10000).then(element=>{element.click()})
  await new Promise(resolve => setTimeout(resolve, 5000));
}
async function navReportPerfSumm(){
  //reporte 2
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[3]/td[2]/h2')).click()
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[4]/td/div/table/tbody/tr[12]/td[2]/a')).click()
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[3]/td[2]/h2'))),10000).then(element=>{element.click()})
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[4]/td/div/table/tbody/tr[12]/td[2]/a'))),10000).then(element=>{element.click()})
  await new Promise(resolve => setTimeout(resolve, 5000));
}
async function navReportDailyOps(){
//reporte 3
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[1]/td[2]/h2')).click()
  // await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td[2]/a')).click()
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[1]/td[2]/h2'))),10000).then(element=>{element.click()})
  await driver.wait(until.elementIsVisible(driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td[2]/a'))),10000).then(element=>{element.click()})
  await new Promise(resolve => setTimeout(resolve, 5000));
}
async function navHome(){
  await driver.get(ipadd+'mainPortal.jsp');
  await driver.wait(until.elementLocated(By.css('#sideMenu')),10000)
  // await new Promise(resolve => setTimeout(resolve, 7000));
  const frameSideMenu=driver.findElement(By.css('#sideMenu'))
  await driver.switchTo().frame(frameSideMenu)
  await driver.wait(until.elementLocated(By.xpath('/html/body/div[3]/div[2]/a[4]')),10000).then(element=>{element.click()})
  //  await driver.findElement(By.xpath('/html/body/div[3]/div[2]/a[4]')).click()
  await new Promise(resolve => setTimeout(resolve, 10000));
  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
}
async function navWkReport(){
  //reporte semanal
  console.log("Entrada a navegacion por semana")
  let actions = driver.actions();
  await driver.wait(until.elementLocated(By.id('revenueCenterData')),10000)
  const selectRevenue = await driver.findElement(By.id('revenueCenterData'))
  select = new Select(selectRevenue)
  await select.selectByValue('23702')
// await driver.findElement(By.id('calendarBtn')).click()
  await driver.wait(until.elementLocated(By.id('calendarBtn')),10000).then(element=>{element.click()})
  await driver.wait(until.elementIsVisible(driver.findElement(By.id('calendarFrame'))),10000)
  const frameCalendar=driver.findElement(By.id('calendarFrame'))
  await driver.switchTo().frame(frameCalendar)

  const selectDates = await driver.findElement(By.id('selectQuick'))
  select= new Select(selectDates)
  await select.selectByValue('Past7Days')

  /*
  var startWeek = new Date();
  startWeek.setDate(startWeek.getDate()-7)      
  var endWeek = new Date();
  endWeek.setDate(endWeek.getDate()-1)
  var middleWeek = new Date();
  middleWeek.setDate(middleWeek.getDate()-6)  
  var row=((startWeek.getMonth()+1)/4)
  var col=((startWeek.getMonth()+1)%4)
  var day=startWeek.getDate()
  let mes      
  if(row<1.25){
    mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[2]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))        
  }else{
    if(row<2.25){
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[3]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }else{
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[4]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }
  }      
  for(let i=0;i<mes.length;i++){
    let aux= await mes[i].getCssValue("background-color")
    if(aux=="rgb(239, 239, 239)"){
      mes.splice(i,1)
    }
  }  
  await actions.click(mes[0]).perform()
  let actions1 = driver.actions();
  row=((middleWeek.getMonth()+1)/4)
  col=((middleWeek.getMonth()+1)%4)
  day=middleWeek.getDate()        
  if(row<1.25){
    mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[2]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))        
  }else{
    if(row<2.25){
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[3]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }else{
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[4]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }
  }      
  for(let i=0;i<mes.length;i++){
    let aux= await mes[i].getCssValue("background-color")
    if(aux=="rgb(239, 239, 239)"){
      mes.splice(i,1)
    }
  }  
  await actions1.keyDown(Key.SHIFT).click(mes[0]).perform().then(element=>{console.log(element)})
  let actions2 = driver.actions();
  row=((endWeek.getMonth()+1)/4)
  col=((endWeek.getMonth()+1)%4)
  day=endWeek.getDate()      
  if(row<1.25){
    mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[2]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))        
  }else{
    if(row<2.25){
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[3]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }else{
      mes= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[4]/div['+col.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+day.toString()+"']"))
    }
  }      
  for(let i=0;i<mes.length;i++){
    let aux= await mes[i].getCssValue("background-color")
    if(aux=="rgb(239, 239, 239)"){
      mes.splice(i,1)
    }
  }  
  await actions2.keyDown(Key.SHIFT).click(mes[0]).keyUp(Key.SHIFT).perform().then(element=>{console.log(element)})
  */


  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
  await driver.findElement(By.id('Run Report')).click()

  await new Promise(resolve => setTimeout(resolve, 7000));

  await driver.wait(until.elementLocated(By.id('excelButton')),7000).then(element=>{element.click()})
  // await driver.findElement(By.id('excelButton')).click()
  await new Promise(resolve => setTimeout(resolve, 10000));
}

async function navDailyReport(){
  console.log("entrada a navegacion diaria")
   // reporte de dia anterior
  await driver.wait(until.elementLocated(By.id('revenueCenterData')),10000)
  const selectRevenue = await driver.findElement(By.id('revenueCenterData'))
  select = new Select(selectRevenue)
  await select.selectByValue('23702')

  const selectDates = await driver.findElement(By.id('calendarData'))
  select = new Select(selectDates )
  await select.selectByValue('Yesterday')
  // const selectRevenue = await driver.findElement(By.id('revenueCenterData'))

  await driver.wait(until.elementLocated(By.id('calendarBtn')),10000).then(element=>{element.click()})
  // await driver.findElement(By.id('calendarBtn')).click()

  await driver.wait(until.elementIsVisible(driver.findElement(By.id('calendarFrame'))),10000)
  const frameCalendar=driver.findElement(By.id('calendarFrame'))
  await driver.switchTo().frame(frameCalendar)

  /*
  var dayAnt= new Date();
  dayAnt.setDate(dayAnt.getDate()-1)
  var rowA=((dayAnt.getMonth()+1)/4)
  var colA=((dayAnt.getMonth()+1)%4)
  var dayA=dayAnt.getDate()
  let mesA      
  if(rowA<1.25){
    mesA= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[2]/div['+colA.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+dayA.toString()+"']"))        
  }else{
    if(rowA<2.25){
      mesA= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[3]/div['+colA.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+dayA.toString()+"']"))
    }else{
      mesA= await (await driver.findElement(By.xpath('/html/body/table/tbody/tr[2]/td[3]/div[3]/div[1]/div/div[4]/div['+colA.toString()+']'))).findElements(By.xpath("table/tbody/tr/td/a[text()='"+dayA.toString()+"']"))
    }
  }      
  for(let i=0;i<mesA.length;i++){
    let aux= await mesA[i].getCssValue("background-color")
    if(aux=="rgb(239, 239, 239)"){
      mesA.splice(i,1)
    }
  }  
  await mesA[0].click()
 */


  await driver.switchTo().defaultContent();
  const frameMyPageA=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPageA)

  
  await driver.findElement(By.id('Run Report')).click()

  await new Promise(resolve => setTimeout(resolve, 7000));    

  await driver.wait(until.elementLocated(By.id('excelButton')),7000).then(element=>{element.click()})
  //  await new Promise(resolve => setTimeout(resolve, 2000));
  // await driver.findElement(By.id('excelButton')).click()
  await new Promise(resolve => setTimeout(resolve, 10000));
}

async function generateReports() {
  while(tries>0 && !correct){
    console.log("Intento "+tries+" en ejecucion")
    try {
    await driver.get(ipadd);
    await navLogin()
    let today=new Date()    
    if(today.getDay()==1){
      console.log("Caso inicio de semana \n")
      await navReportDailyOps()
      await navWkReport()
      await navHome()
      await transformReportDailyOps(true)
      console.log("Fin reporte semanal DailyOps \n")
      await navReportPerfSumm()
      await navWkReport()
      await navHome()
      await transformReportDailyServPerf(true)
      console.log("Fin reporte semanal ServPerf \n")
      await navReportMixItems()
      await navWkReport()
      await transformReportDailyMixItem(true)
      console.log("Fin reporte semanal MixItem \n")
      await navHome()
      await navReportDailyOps()
      await navDailyReport()
      await navHome()
      await transformReportDailyOps(false)
      console.log("Fin reporte diario DailyOps \n")
      await navReportPerfSumm()
      await navDailyReport()
      await navHome()
      await transformReportDailyServPerf(false)
      console.log("Fin reporte diario ServPerfDaily \n")
      await navReportMixItems()
      await navDailyReport()
      await transformReportDailyMixItem(false)
      console.log("Fin reporte diario MixItem \n")
    }else{
      console.log("Caso intermedio de semana")
      await navReportDailyOps()
      await navDailyReport()
      await navHome()
      await transformReportDailyOps(false)
      console.log("Fin reporte diario DailyOps \n")
      await navReportPerfSumm()
      await navDailyReport()
      await navHome()
      await transformReportDailyServPerf(false)
      console.log("Fin reporte diario ServPerfDaily \n")
      await navReportMixItems()
      await navDailyReport()
      await transformReportDailyMixItem(false)
      console.log("Fin reporte diario MixItem")
    }
    console.log("Intento "+tries+" exitoso")
    correct=true
  } catch(err){
    console.log("Intento "+tries+" fallido")
    tries-=1
    emptyFolder()
    console.error(err)
  }finally {
    driver.quit();
  }
  }
  
};

async function transformReportDailyOps(rptwkn){
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'DailyOpsReport.xlsx');  
  let page= await doc.getWorksheet(1)
  let colB=page.getColumn(2).values
  let colF=page.getColumn(6).values
  let dates,numDate,date1,date2,date
  if(rptwkn){
    dates=colB[2].split('-')
    console.log(dates)
    numDate=dates[0].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    numDate=dates[1].split('/')
    date2=numDate[1]+'/'+numDate[0].trim()+'/'+numDate[2].substring(0,4)
    date=date1+'-'+date2    
    await changeFileName(downloadFilepath+'DailyOpsReport',date)
  }else{
    numDate=colB[2].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    date=date1
    await changeFileName(downloadFilepath+'DailyOpsReport',date)
  }
  const result=new ExcelJS.Workbook();
  let pagina=result.addWorksheet()
  let cod=parseInt(colB[4].substring(0,colB[4].indexOf(' ')))
  let revenue=colB[4].substring(colB[4].indexOf(' ')+1,colB[4].length)
  let disc=colB[9]
  pagina.addRow(['Site ID','Site Name','Date of Business','Sales Channel','Total Net Sales','Total Tax','Total Gross Sales','Total Comps','Total Discounts','Total Transactions','Total Customers'])  
  pagina.addRow([cod,revenue,date,'Dine In',colB[7],colB[13],colB[8],0,disc,colF[9],colF[8]])
  pagina.getCell('E2').numFmt='#,##0.00'
  pagina.getCell('F2').numFmt='#,##0.00'
  pagina.getCell('G2').numFmt='#,##0.00'
  pagina.getCell('I2').numFmt='#,##0.00'
  resizeColumn(pagina)
  await result.xlsx.writeFile(downloadFilepath+'Sales.xlsx')
  let docname= changeFileName(downloadFilepath+'Sales',date)  
  await uploadFileSFTP(docname)
}
async function transformReportDailyServPerf(rptwkn){
  let index=2
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'ServiceDailyDetail.xlsx');
  let page= await doc.getWorksheet(1)
  let colA=page.getColumn(1).values
  let colB=page.getColumn(2).values
  let dates,numDate,date1,date2,date
  if(rptwkn){
    dates=colB[2].split('-')
    numDate=dates[0].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    numDate=dates[1].split('/')
    date2=numDate[1]+'/'+numDate[0].trim()+'/'+numDate[2].substring(0,4)
    date=date1+'-'+date2
    await changeFileName(downloadFilepath+'ServiceDailyDetail',date)
  }else{
    numDate=colB[2].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    date=date1
    await changeFileName(downloadFilepath+'ServiceDailyDetail',date)
  }
  let cod=parseInt(colB[4].substring(0,colB[4].indexOf(' ')))
  let revenue=colB[4].substring(colB[4].indexOf(' ')+1,colB[4].length)
  const result=new ExcelJS.Workbook();
  let pagina=result.addWorksheet()
  const init = (element) => element == 'Hour';
  const end = (element) => element == 'All Fixed Periods';
  const IndInit=colA.findIndex(init)
  const Indend=colA.findIndex(end)  
  pagina.addRow(['Site ID','Site Name','Business Date','Start Time','End Time','Total Sales','Total Customers','Total Transactions'])  
  for(let i=IndInit+1;i<Indend;i++){
    let aux=page.getRow(i).values
    pagina.addRow([cod,revenue,date,aux[1],aux[1].replace(":00",":59"),aux[2],aux[6],aux[8]])
    pagina.getCell('F'+index.toString()).numFmt="#,##0.00"
    index+=1
  }
  resizeColumn(pagina)  
  await result.xlsx.writeFile(downloadFilepath+'Salesbydaypart.xlsx')
  let docname= changeFileName(downloadFilepath+'Salesbydaypart',date)  
  await uploadFileSFTP(docname)
}
async function transformReportDailyMixItem(rptwkn){
  let index=2
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'SalesMixItemsSummary.xlsx');
  let page= await doc.getWorksheet(1)
  let colA=page.getColumn(1).values
  let colB=page.getColumn(2).values
  let dates,numDate,date1,date2,date
  if(rptwkn){
    dates=colB[2].split('-')
    numDate=dates[0].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    numDate=dates[1].split('/')
    date2=numDate[1]+'/'+numDate[0].trim()+'/'+numDate[2].substring(0,4)
    date=date1+'-'+date2
    await changeFileName(downloadFilepath+'SalesMixItemsSummary',date)
  }else{
    numDate=colB[2].split('/')
    date1=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
    date=date1
    await changeFileName(downloadFilepath+'SalesMixItemsSummary',date)
  }
  let cod=parseInt(colB[4].substring(0,colB[4].indexOf(' ')))
  let revenue=colB[4].substring(colB[4].indexOf(' ')+1,colB[4].length)
  const result=new ExcelJS.Workbook();
  let pagina=result.addWorksheet()
  const init = (element) => element == 'Item';
  const end = (element) => element == 'Total Item Sales:';
  const IndInit=colA.findIndex(init)
  const Indend=colA.findIndex(end)  
  pagina.addRow(['Site ID','Site Name','Business Date','Item Number','Menu Item Name','Menu Item Category 1','Menu Item Category 2','Menu Item Category 3','Total Qty Sold','Selling Price','Discounted Price','Net Sales'])  
  for(let i=IndInit+1;i<Indend;i++){
    let aux=page.getRow(i).values
    if(aux[4]>0){
      pagina.addRow([cod,revenue,date,aux[7],aux[1],aux[3],aux[2],"",aux[8],aux[4]/aux[8],aux[5],aux[6]])
      pagina.getCell('J'+index.toString()).numFmt="#,##0.00"
      pagina.getCell('K'+index.toString()).numFmt="#,##0.00"
      pagina.getCell('L'+index.toString()).numFmt="#,##0.00"
      index+=1
    }    
  }
  resizeColumn(pagina)
  await result.xlsx.writeFile(downloadFilepath+'MenuMix.xlsx')
  let docname= changeFileName(downloadFilepath+'MenuMix',date)
  console.log(docname)
  await uploadFileSFTP(docname)
}
async function uploadFileSFTP(filename) {
  let name=filename.split("\\")[4];
  const sftp = new Client();
  try {    
    await sftp.connect({
      host: 'ftp.fogodechao.com',
      port: 22,
      username: 'fogoecuador',
      password: 'E6y57Wr6%wd^!jagqJ@M'
    });
    await sftp.put(filename, "/"+name);
    console.log(`Archivo ${filename} subido correctamente`);
    console.log(`Archivo ${name} subido correctamente`);
  } catch (error) {
    console.error(`Error al subir el archivo: ${error.message}`);
  } finally {    
    sftp.end();
  }
}
async function resizeColumn(worksheet){
  worksheet.columns.forEach(function (column, i) {
    let maxLength = 0;
    column["eachCell"]({ includeEmpty: true }, function (cell) {
        var columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength ) {
            maxLength = columnLength;
        }
    });
    column.width = maxLength < 10 ? 10 : (maxLength+5);
});
}
function changeFileName(filename,date){  
  let name = date.replace("-","_to_")
  name=name.replace("/","-")
  name=name.replace("/","-")
  name=name.replace("/","-")
  name=name.replace("/","-")
  fs.rename(filename+".xlsx", filename+"_"+name+".xlsx", (err) => {    
    if (err) {
      console.error(`Error al cambiar el nombre del archivo ${filename+".xlsx"}:`,err);
    } else {
      console.log(`Archivo ${filename+".xlsx"} renombrado como: ${filename+"_"+name+".xlsx"}`);
    }
  });
  return filename+"_"+name+".xlsx"
}

async function emptyFolder(){
  const archivos = fs.readdirSync(downloadFilepath);
  archivos.forEach((archivo) => {
    const rutaArchivo = path.join(downloadFilepath, archivo);
    fs.unlinkSync(rutaArchivo);    
  });
}

generateReports()

