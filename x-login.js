const pBrowser = require("../../utils/pBrowser.js");
const twitter = require("./xUtils.js");
const path = require("path");
const fs = require("fs");
const { sleep } = require("../../utils/time.js");
const { getExcelDataList, updateOneDataInExcel } = require("../../utils/ExcelDataUtils");
const file = "../../data/x.xlsx";
var excelDataList = [];

async function getExcelData() {
    excelDataList = await getExcelDataList(file);
    if (excelDataList.length == 0) {
        console.log(`请配置 ${file} `);
        return [];
    }
    
    const updatedArray = excelDataList.filter(object => object.xusername && object.success != true);
    console.log(`待登录的账号数: ${updatedArray.length} 条`);
    return updatedArray;
}

async function openBrowser(itemData) {
    console.log(`[${new Date().toLocaleString()}] 步骤2: 正在打开浏览器窗口 ${itemData.windowName}...`);
    try {
        // if (!global.config || !global.config.browsers) {
        //     global.config = { browsers: [] };
        // }
        
        console.log("windowName", itemData.windowName);
        const browser = await pBrowser.getBrowserByName(itemData.windowName);
        
        if (browser) {
            const browserObj = { [itemData.windowName]: browser };
            global.config.browsers.push(browserObj);
            console.log(`打开 ${itemData.windowName} 成功`);
            return true;
        } else {
            console.error(`打开 ${itemData.windowName} 失败`);
            return false;
        }
    } catch (err) {
        console.error(err);
        return false;
    }
}

async function loginX(itemData) {
    try {
        console.log("== windowName = ",itemData.windowName);
        const res = await twitter.login(itemData);
        if(res){
            itemData.success = true;
            await updateOneDataInExcel(excelDataList, itemData.id, itemData, file);
            await sleep(500);
            console.log(` ${itemData.windowName}  login success `);
            return true;
        }else{
            console.log(` ${itemData.windowName}  login false  `);
            return false;
        }
    }catch(err){
        console.error(` ${itemData.windowName}  login err: ${err}  `);
        return false;
    }
      
}

async function closeAllBrowser() {
    global.config.browsers.forEach(object => {
        for (let key in object) {
            const bro = object[key];
            bro.close();
            console.log(`winName ${key} 关闭窗口成功`);
        }
    });
    global.config.browsers = [];
}

async function main() {
    let t = 0;
    const cols = 4;
    while (t < 2) {
        const dataList = await getExcelData();
        if (dataList.length == 0) {
            console.log("所有账号已登录，任务结束。");
            process.exit(0);
        }
        
        global.config = {};
        global.config.browsers = [];
        global.config.userList = await pBrowser.getAllUserList(); 
        await sleep(2000);
        console.log("========================开始============================");
        // console.log(global.config.userList);
        
        for (let i = 0; i < dataList.length; i += cols) { 
            let data_items = [];
            for (let j = i; j < i + cols && j < dataList.length; j++) { 
                let itemData = dataList[j]; // Modify the value as an example 
                if( itemData.xusername && itemData.xpassword && itemData.success != true ){
                    data_items.push(itemData)
                    await openBrowser( itemData );
                }
                // await sleep(500)
            }
            await sleep(500);
            let res = await pBrowser.reSizeWindow_3();
            await sleep(500);

            let tasks = data_items.map(( iData ) => loginX( iData ));
            await Promise.all(tasks); 
            await sleep(10000);

            await closeAllBrowser();
            await sleep(1000);
        } // Update the table with the new data
        t++;
    } 
}

main();
