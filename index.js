const {Builder, By, Key, until,Select,WebElement,Actions} = require('selenium-webdriver');
const firefox = require('selenium-webdriver/firefox');
const ExcelJS = require('exceljs');
const options= new firefox.Options()
const downloadFilepath = "C:\\RPA\\reports\\"
const ipadd = "http://10.163.132.13/"
// options.addArguments('-headless')
// options.addArguments('-private')
options.setPreference("browser.download.folderList",2)
options.setPreference("browser.download.manager.showWhenStarting",false)
options.setPreference("browser.download.dir",downloadFilepath)
options.setPreference("browser.helperApps.neverAsk.saveToDisk",'application/xlsx')
let driver =  new Builder().forBrowser('firefox').setFirefoxOptions(options).build();
const actions = driver.actions();
const fs = require('fs');
const path = require('path');
let Client = require('ssh2-sftp-client');
const sftphost = 'test.rebex.net';
const sftpport = 22;
const sftpusername = 'demo';
const sftppassword = 'password';
async function navLogin(){
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[3]')).sendKeys('rpafogo');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[4]')).sendKeys('mmh');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/input[5]')).sendKeys('Hotel321#');
  await driver.findElement(By.xpath('/html/body/div[1]/div[1]/div[5]/div/form/button')).click();
  await new Promise(resolve => setTimeout(resolve, 6000));
  const frameSideMenu=driver.findElement(By.css('#sideMenu'))
  await driver.switchTo().frame(frameSideMenu)
  await driver.findElement(By.xpath('/html/body/div[3]/div[2]/a[4]')).click()
  await new Promise(resolve => setTimeout(resolve, 2000));
  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
}

async function navReportMixItems(){
  //reporte 1 
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[7]/td[2]/h2')).click()
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[8]/td/div/table/tbody/tr[6]/td[2]/a')).click()
  await new Promise(resolve => setTimeout(resolve, 3000));
}

async function navReportPerfSumm(){
  //reporte 2
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[3]/td[2]/h2')).click()
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[4]/td/div/table/tbody/tr[12]/td[2]/a')).click()
  await new Promise(resolve => setTimeout(resolve, 3000));
}

async function navReportDailyOps(){
//reporte 3
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[1]/td[2]/h2')).click()
  await driver.findElement(By.xpath('/html/body/p[2]/table/tbody/tr[2]/td/div/table/tbody/tr[1]/td[2]/a')).click()
  await new Promise(resolve => setTimeout(resolve, 3000));
}

async function navHome(){
  await driver.get(ipadd+'mainPortal.jsp');
  await new Promise(resolve => setTimeout(resolve, 6000));
  const frameSideMenu=driver.findElement(By.css('#sideMenu'))
  await driver.switchTo().frame(frameSideMenu)
  await driver.findElement(By.xpath('/html/body/div[3]/div[2]/a[4]')).click()
  await new Promise(resolve => setTimeout(resolve, 2000));
  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
}

async function navWkReport(){
  //reporte semanal
  const selectRevenue = await driver.findElement(By.id('revenueCenterData'))
  select = new Select(selectRevenue)
  await select.selectByValue('23702')
  await driver.findElement(By.id('calendarBtn')).click()
  const frameCalendar=driver.findElement(By.id('calendarFrame'))
  await driver.switchTo().frame(frameCalendar)
  var startWeek = new Date();
  startWeek.setDate(startWeek.getDate()-7)      
  var endWeek = new Date();
  endWeek.setDate(endWeek.getDate()-1)  
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
  // driver.executeScript("arguments[0].setAttribute('style','background-color:#f2c80f')",mes[0])
  await mes[0].click()
  row=((endWeek.getMonth()+1)/4)
  col=((endWeek.getMonth()+1)%4)
  day=endWeek.getDate()
  mes      
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
  // driver.executeScript("arguments[0].setAttribute('style','background-color:#f2c80f')",mes[0])
  await actions.keyDown(Key.SHIFT).click(mes[0]).keyUp(Key.SHIFT).perform()
  await driver.switchTo().defaultContent();
  const frameMyPage=driver.findElement(By.xpath('//*[@id="myPage"]'))
  await driver.switchTo().frame(frameMyPage)
  await driver.findElement(By.id('Run Report')).click()
  await new Promise(resolve => setTimeout(resolve, 2000));
  await driver.findElement(By.id('excelButton')).click()
  await new Promise(resolve => setTimeout(resolve, 10000));
}

async function navDailyReport(filename){
   // reporte de dia anterior
  const selectDate = await driver.findElement(By.id('calendarData'))
  var select = new Select(selectDate)
  await select.selectByValue('Yesterday')
  const selectRevenue = await driver.findElement(By.id('revenueCenterData'))
  select = new Select(selectRevenue)
  await select.selectByValue('23702')
  await driver.findElement(By.id('Run Report')).click()
  await new Promise(resolve => setTimeout(resolve, 2000));
  await driver.findElement(By.id('excelButton')).click()
  await new Promise(resolve => setTimeout(resolve, 10000));
}

async function generateReports() {  
  try {    
    await driver.get(ipadd);
    await navLogin()
    let today=new Date()    
    if(today.getDay()==1){
      await navReportDailyOps()
      await navWkReport()
      await navHome()
      await transformReportDailyOps()
      await navReportPerfSumm()
      await navWkReport()
      await navHome()
      await transformReportDailyServPerf()
      await navReportMixItems()
      await navWkReport()
      await transformReportDailyMixItem()
      await navHome()
      await navReportDailyOps()
      await navDailyReport()
      await navHome()
      await transformReportDailyOps()
      await navReportPerfSumm()
      await navDailyReport()
      await navHome()
      await transformReportDailyServPerf()
      await navReportMixItems()
      await navDailyReport()
      await transformReportDailyMixItem()
    }else{
      await navReportDailyOps()
      await navDailyReport()
      await navHome()
      await transformReportDailyOps()
      await navReportPerfSumm()
      await navDailyReport()
      await navHome()
      await transformReportDailyServPerf()
      await navReportMixItems()
      await navDailyReport()
      await transformReportDailyMixItem()
    }
  } catch(err){
    console.log(err)
  }finally {
    driver.quit();
  }
};

async function transformReportDailyOps(){
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'DailyOpsReport.xlsx');  
  let page= await doc.getWorksheet(1)
  let colB=page.getColumn(2).values
  let colF=page.getColumn(6).values
  let numDate=colB[2].split('/')  
  let date=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
  const result=new ExcelJS.Workbook();
  let pagina=result.addWorksheet()
  let cod=colB[4].substring(0,colB[4].indexOf(' '))
  let revenue=colB[4].substring(colB[4].indexOf(' ')+1,colB[4].length)
  let disc=Math.abs(colB[9])  
  pagina.addRow(['Site ID','Site Name','Date of Business','Sales Channel','Total Net Sales','Total Tax','Total Gross Sales','Total Comps','Total Discounts','Total Transactions','Total Customers'])
  pagina.addRow([cod,revenue,date,'Dine In',formatNumber(colB[7]),formatNumber(colB[13]),formatNumber(colB[8]),0,formatNumber(disc),colF[9],colF[8]])  
  await result.xlsx.writeFile(downloadFilepath+'DailyOpsFrm.xlsx')
}
async function transformReportDailyServPerf(){
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'ServiceDailyDetail.xlsx');
  let page= await doc.getWorksheet(1)
  let colA=page.getColumn(1).values
  let colB=page.getColumn(2).values
  let numDate=colB[2].split('/')  
  let date=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
  let cod=colB[4].substring(0,colB[4].indexOf(' '))
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
    pagina.addRow([cod,revenue,date,aux[1],aux[1].replace(":00",":59"),formatNumber(Math.abs(aux[2])),aux[6],aux[8]])
  }
  await result.xlsx.writeFile(downloadFilepath+'ServiceDailyDetailFrm.xlsx')
}
async function transformReportDailyMixItem(){
  const workbook = new ExcelJS.Workbook();
  let doc= await workbook.xlsx.readFile(downloadFilepath+'SalesMixItemsSummary.xlsx');
  let page= await doc.getWorksheet(1)
  let colA=page.getColumn(1).values
  let colB=page.getColumn(2).values
  let numDate=colB[2].split('/')  
  let date=numDate[1]+'/'+numDate[0]+'/'+numDate[2].substring(0,4)
  let cod=colB[4].substring(0,colB[4].indexOf(' '))
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
    pagina.addRow([cod,revenue,date,aux[7],aux[1],aux[3],aux[2],"",aux[8],formatNumber(Math.abs(aux[4])/Math.abs(aux[8])),formatNumber(Math.abs(aux[5])),formatNumber(Math.abs(aux[6]))])
  }
  await result.xlsx.writeFile(downloadFilepath+'SalesMixItemsSummaryFrm.xlsx')
}

function formatNumber(number) {  
  if (isNaN(number)) {
    return "No es un número válido";
  }
  const numbRound = Number(number).toFixed(2);  
  const parts = numbRound.toString().split(".");  
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  const numFormato = parts.join(".");
  return numFormato;
}

function emptyFolder(ruta) {
  if (fs.existsSync(ruta)) {
    fs.readdirSync(ruta).forEach((archivo, index) => {
      const archivoRuta = path.join(ruta, archivo);
      if (fs.lstatSync(archivoRuta).isDirectory()) {
        emptyFolder(archivoRuta);
      } else {
        fs.unlinkSync(archivoRuta);
      }
    });    
    console.log(`Carpeta "${ruta}" vaciada exitosamente.`);
  } else {
    console.log(`La carpeta "${ruta}" no existe.`);
  }
}
async function uploadFileSFTP(filename) {
  const sftp = new Client();
  try {    
    await sftp.connect({
      host: '',
      port: 22,
      username: '',
      password: ''
    });
    await sftp.put(downloadFilepath+filename, "/"+filename);
    console.error(`Archivo ${filename} subido correctamente`);
  } catch (error) {
    console.error(`Error al subir el archivo: ${error.message}`);
  } finally {    
    sftp.end();
  }
}
emptyFolder(downloadFilepath)
generateReports()