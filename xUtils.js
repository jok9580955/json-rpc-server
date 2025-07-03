const { sleep } = require("../../utils/time.js");
const pBrowser= require("../../utils/pBrowser.js");
const {getRandomInt} = require("../../utils/Random.js");
const dmail= require("../../utils/dmail.js");
const { shuffleArray } = require("../../utils/array.js");
const {getExcelDataList, updateOneDataInExcel} = require("../../utils/ExcelDataUtils");
const path = require("path");
const fs = require("fs");
const file = "./tweets.xlsx";
const tweetUrlFile = "./tweetsurl.xlsx";
var excelDataList = [];

const xUrl = 'https://x.com/home';
const xPre = 'https://x.com/';
const xLoginUrl = 'https://x.com/i/flow/login';

const X_confirm = "Your X confirmation code is";
const X_access = "confirm your email address to access";

const loginAgain = async (page,itemData) => {
    if(!itemData){
        console.log("请配置数据表格");
        return false;
    }
    if(!page){
        console.log(" page===null ");
        return false;
    }
    const password = itemData.xpassword;
    const username = itemData.xusername;
    try{     
        let nameSelector = 'input[autocomplete="username"]';
        await page.waitForSelector(nameSelector,{timeout: 3000});
        await page.focus(nameSelector);
        await sleep(200);
        await page.keyboard.down('Control');
        await sleep(200);
        await page.keyboard.press('A');
        await sleep(200);
        await page.keyboard.up('Control');
        await sleep(200);
        await page.keyboard.press("Backspace");
        await sleep(200);
        await page.keyboard.type(username); 
        await sleep(200);
        await page.waitForSelector(`::-p-text(Next)`);
        await page.click(`::-p-text(Next)`);
        await sleep(3000);
        let passSelector = 'input[name="password"]'; //autocomplete="current-password"       
        await page.waitForSelector(passSelector,{timeout: 2000});
        await page.focus(passSelector);
        await sleep(200);
        await page.keyboard.down('Control');
        await sleep(200);
        await page.keyboard.press('A');
        await sleep(200);
        await page.keyboard.up('Control');
        await sleep(200);
        await page.keyboard.press("Backspace");
        await sleep(200);
        await page.keyboard.type(password); 
        await sleep(100);
        // await page.keyboard.press("Tab");
        // await sleep(200);
        await page.keyboard.press("Enter");
        await sleep(2000);
        return true;
    }catch(err ){
        console.log(' no selector ============== input[autocomplete="username"] ');
        return false;
    }
    
}
/*
*  登录
*/
const login = async (walletData) => {
    if(!walletData ){
        console.log("请配置数据表格");
        return false;
    }
    const browser = await pBrowser.getBrowserByGlobalConfig(walletData.windowName); //global.config.browser;
    // console.log("loginOkx  browser : ",browser);
    if(!browser)
        return false
    const password = walletData.xpassword;
    const username = walletData.xusername;
    const privateKey = walletData.privateKey;
    //打开okx插件
    await sleep(3000);
    for(let i=0;i<3;i++)
    {
        const pages = await browser.pages();
        const pagelen = pages.length;
        if(pagelen>1){
            for(let j=pagelen-1; j>0; j--){
                let _page = pages[j];
                if(_page){
                    await _page.close();
                }
                await sleep(500);
            }
            await sleep(3000);
        }    
    }
    //打开页面
    let page = await browser.newPage();
    try{
   
        await page.goto(xUrl, { waitUntil:"load", timeout:30000 });//networkidle0
        await sleep(3000);
        // await page.waitForNavigation();
        // await page.waitForNetworkIdle();
        await page.bringToFront();
       
    }catch(err){
        console.log("err:",err);
    }
    let res = await checkRetry(page);
    if(!res){
        return false;
    }
    // if(page.url().includes('x.com/home'))
    // {
    //     console.log("  ===已经登录着=== x.com/home======= ");
            
    //     return true;
    // }else{
    if(page.url().includes('x.com/account/access'))
    {
        if(!privateKey){
            console.log(" 缺少 privateKey 无法获取 dmail 中的验证码 ");
            return false;
        }
        res = await checkAccess(page, privateKey);
    }
    // }
    try{     
        let nameSelector = 'input[autocomplete="username"]';
        await page.waitForSelector(nameSelector,{timeout: 3000});
        await page.focus(nameSelector);
        await sleep(200);
        await page.keyboard.down('Control');
        await sleep(200);
        await page.keyboard.press('A');
        await sleep(200);
        await page.keyboard.up('Control');
        await sleep(200);
        await page.keyboard.press("Backspace");
        await sleep(200);
        await page.keyboard.type(username); 
        await sleep(200);
        await page.waitForSelector(`::-p-text(Next)`);
        await page.click(`::-p-text(Next)`);
        await sleep(3000);
        let passSelector = 'input[name="password"]'; //autocomplete="current-password"       
        await page.waitForSelector(passSelector,{timeout: 2000});
        await page.focus(passSelector);
        await sleep(200);
        await page.keyboard.down('Control');
        await sleep(200);
        await page.keyboard.press('A');
        await sleep(200);
        await page.keyboard.up('Control');
        await sleep(200);
        await page.keyboard.press("Backspace");
        await sleep(200);
        await page.keyboard.type(password); 
        await sleep(100);
        // await page.keyboard.press("Tab");
        // await sleep(200);
        await page.keyboard.press("Enter");
        await sleep(8000);
        try{
            if(page.url().includes('x.com/account/access'))
            {
                if(!privateKey){
                    console.log(" 缺少 privateKey 无法获取 dmail 中的验证码 ");
                    return false;
                }
                await checkAccess(page, privateKey);
            }else{
                await page.waitForSelector('input[data-testid="ocfEnterTextTextInput"]',{timeout: 5000});
                await page.focus('input[data-testid="ocfEnterTextTextInput"]');
                await sleep(1000);
                if(!privateKey){
                    console.log(" 缺少 privateKey 无法获取 dmail 中的验证码 ");
                    return false;
                } 
                const verifyCode = await getEmailCode(privateKey);
                if(!verifyCode){
                    console.log(" 拿不到验证码。。。 ");
                    return false;
                }
                await page.keyboard.type(verifyCode);
                await sleep(3000);
                await page.bringToFront();
                await page.keyboard.press("Enter");
                await sleep(3000);
                
            }     
        }catch(err){

        } 
    }catch(err ){
        console.log(' no selector ============== input[autocomplete="username"] ');
      
    }
    try{
        await page.waitForSelector(`::-p-text(Keep less relevant ads)`,{timeout:8000});
        await page.click(`::-p-text(Keep less relevant ads)`);
        await sleep(2000);
    }catch(err){

    }
    try{
        await page.waitForSelector(`::-p-text(Yes, that's my email)`,{timeout:8000});
        await page.click(`::-p-text(Yes, that's my email)`);
        await sleep(2000);
    }catch(err){

    }
    if(page.url().includes('x.com/home'))
    {
        // await sleep(3000);
        // const randomOffset = await getRandomInt(150,360);
        // await browser.scrollPage(page, randomOffset);
        await sleep( getRandomInt(5000,8000) );
        return true;
    }
    return false;
}

const getEmailAccess = async ( privateKey ) =>{
    // const regex = /^\s\d{6}\s$/;
    // const code = text.catch(regex);
    const tokenInfo = await dmail.getDmailTokenInfoByPirvateKey(privateKey);
    let params = {};
    params.token = tokenInfo.token;
    params.pid = tokenInfo.pid;
    params.type = X_access;
    // const X_confirm = "confirm your email address to access";
    await sleep(2000);
    const confirmCode = await dmail.getAccessCode(params);
    return confirmCode?confirmCode:null;
}

const getEmailCode = async ( privateKey ) => {
    const tokenInfo = await dmail.getDmailTokenInfoByPirvateKey(privateKey);
    let params = {};
    params.token = tokenInfo.token;
    params.pid = tokenInfo.pid;
    params.type = X_confirm;
    // const X_confirm = "Your X confirmation code is";
    await sleep(2000);
    const confirmCode = await dmail.getVerifyCode(params);
    return confirmCode?confirmCode:null;

}

const follow = async ( itemData ) => {
    if(!itemData ){
        console.log("请配置数据表格");
        return false;
    }
    const browser = await pBrowser.getBrowserByGlobalConfig(itemData.windowName); //global.config.browser;
    // console.log("loginOkx  browser : ",browser);
    if(!browser){
        return false
    }
    await sleep(3000);
    const pages = await browser.pages();
    const page = await pages.find( mpage => mpage.url().includes("x.com")) || null;
    if(!page){
        page = await browser.newPage();
    }
    const followArr = global.config.toFollows[ itemData.windowName ];
    // const followArrTmp = global.config.toFollows;
    // const followArr = shuffleArray(followArrTmp);
    if( !followArr || !(followArr.length)  ){
        return false;
    }
    let succ=0;
    for(let i=0; i<followArr.length; i++){
        let userName = followArr[i];
        let toUrl = xPre+userName;
        console.log(` toUrl ${toUrl}`)
        await page.goto(toUrl,{waitUntil:"load"});
        await sleep(3000);
        await checkRetry(page);
        const aName = '@' + userName;
        const options = {tryTime: 10,
            clickInterval: 1000,
            successSelector: [`[aria-label="Following ${aName}"]`, `[aria-label="正在关注 ${aName}"]`, 
            `::-p-text(Edit profile)`, `::-p-text(编辑个人资料)`
        ]};

        let res = await pBrowser.clickMultipleSelectorsV2(page,
            [ `[aria-label="关注 ${aName}"]`, `[aria-label="Follow ${aName}"]`,
              `[aria-label="回关 ${aName}"]`, `[aria-label="Follow back ${aName}"]`],
            options);
        if(res){
            let followedPath = path.resolve(__dirname, `./followed-txt/${itemData.windowName}.txt`);
            fs.writeFileSync(followedPath, `${userName}\n`, { flag: 'a' });
            console.log(`成功写入文件: ${followedPath}`);
            succ++;
        }

        if(succ == followArr.length )
        {
            return true;
        }
    }
    return false;
}

const checkRetry = async (page) => {
    for(let i=0;i<10;i++){
        try{
            await page.waitForSelector(`::-p-text(Retry)`,{timeout:3000});   
            if(page.url().includes('x.com/home')){
                await page.reload({ waitUntil:"load", timeout:30000 });
            }else{
                await page.click(`::-p-text(Retry)`);    
            }
            await sleep(3000);     
        }catch(err){
            return true;
        }
        
    }
    return false;
}

const checkAccess = async (page, privateKey) => {
    if(page.url().includes('x.com/account/access'))
    {
        if(!privateKey){
            console.log(" 缺少 privateKey 无法获取 dmail 中的验证码 ");
            return false;
        }
        let randomSleep = await getRandomInt( 2,6 )
        await page.waitForSelector(`input[type="submit"][class*="Button"]`,{timeout: 5000});
        await page.click(`input[type="submit"][class*="Button"]`);
        await sleep( randomSleep*1000 );
        await page.waitForSelector(`input[type="submit"][class*="Button"]`,{timeout: 5000});
        await page.click(`input[type="submit"][class*="Button"]`);
        randomSleep = await getRandomInt( 6,12 )
        await sleep( randomSleep*1000 );
        await page.waitForSelector('input[name="token"]',{timeout: 5000});
        await page.focus('input[name="token"]');
        await sleep(1000);
        const verifyCode = await getEmailAccess(privateKey);
        await page.keyboard.type(verifyCode);
        await sleep(1000);
        await page.bringToFront();
        await page.keyboard.press("Enter");
        await sleep(3000);
        await page.waitForSelector(`input[type="submit"][class*="Button"]`,{timeout: 5000});
        await page.click(`input[type="submit"][class*="Button"]`);
        await sleep(3000);
        return true;
    }
    return false;
}

const tweet = async (itemData) => {
    //a[href="/home"]
    if(!itemData ){
        console.log("请配置数据表格");
        return false;
    }
    
    const browser = await pBrowser.getBrowserByGlobalConfig(itemData.windowName); //global.config.browser;
    // console.log("loginOkx  browser : ",browser);
    if(!browser){
        console.log(` can't get browser : ${itemData.windowName} ` );
        return false;
    }
    excelDataList = await getExcelDataList(file);
    if (!Array.isArray(excelDataList) || excelDataList.length === 0) {
        console.log(` 请配置 tweets.xlsx 内容表格 `);
        return false;
    }
    await sleep(3000);
    const pages = await browser.pages();
    const page = await pages.find( mpage => mpage.url().includes("x.com")) || null;
    if(!page){
        // return false;
        page = await browser.newPage();
        await page.goto( xUrl , { waitUntil:"load" });
        await sleep(2000);
    }
    try{
        await page.waitForSelector(`a[href="/home"]`,{timeout:5000});
        await page.click( `a[href="/home"]` );
        await sleep(2000);
        await page.evaluate(() => {
            window.scrollTo(0, 0);
        });
        // await browser.scrollPage(page, );
        await sleep(2000);

        await page.waitForSelector(`.public-DraftStyleDefault-block.public-DraftStyleDefault-ltr`,{timeout:3000});
        await sleep(1000)
        await page.click(`.public-DraftStyleDefault-block.public-DraftStyleDefault-ltr`); // focus
        await sleep(3000);
        const contents = ( await getRamdomContent() ) + " " + ( await getRamdomAt() ) + "  " + ( await getRamdomTopic() ) + " " ;
        await page.keyboard.type(contents);
        await sleep(3000);
        
        // Your post was sent
        const options = {tryTime: 10,
            clickInterval: 1000,
            successSelector: [`::-p-text(Your post was sent)`
        ]};

        let res = await pBrowser.clickMultipleSelectorsV2(page,
            [ `button[data-testid="tweetButtonInline"]`, `button[data-testid="tweetButton"]`],
            options);
        if(res){
            
            await sleep(2000);
            return true;
        }

        console.log(` === no view === post was sent === `);
        return false;

    }catch(err){
        return false;
    }
}

const getRamdomContent = async ( ) => {
    let t = 0;    
    // const randomContent = array[Math.floor(Math.random() * array.length)].content;
    try {
        
        while(t<30){
            const randomItem = excelDataList[Math.floor(Math.random() * excelDataList.length)];
            if(randomItem?.content)
            {
                console.log(` 成功获取随机内容: ${randomItem.content} `);
                return randomItem?.content;
            }
            t++;
            console.log(` try ${t} 次，获取随机内容失败:`);
        }
        // console.log(` try ${t} 次，获取随机内容失败:`);
        return null;
        
    } catch (error) {
        console.error('获取随机内容失败:', error);
        return null;
    }
}

const getRamdomAt = async ( ) => {
    let t = 0;    
    // const randomContent = array[Math.floor(Math.random() * array.length)].content;
    const datalist = await excelDataList.filter( object => object.at );
    try {
        
        while(t<30){
            const randomItem = datalist[Math.floor(Math.random() * datalist.length)];
            if(randomItem?.at)
            {
                console.log(` 成功获取随机 at : ${randomItem.at} `);
                return randomItem?.at;
            }
            t++;
            console.log(` try ${t} 次，获取随机 at 失败:`);
        }
        // console.log(` try ${t} 次，获取随机内容失败:`);
        return null;
        
    } catch (error) {
        console.error(' error  获取随机 at 失败:', error);
        return null;
    }
}

const getRamdomTopic = async ( ) => {
    let t = 0;    
    // const randomContent = array[Math.floor(Math.random() * array.length)].content;
    const datalist = await excelDataList.filter( object => object.topic );
    try {
        
        while(t<30){
            const randomItem = datalist[Math.floor(Math.random() * datalist.length)];
            if(randomItem?.topic)
            {
                console.log(` 成功获取随机内容: ${randomItem.topic} `);
                return randomItem?.topic;
            }
            t++;
            console.log(` try ${t} 次，获取随机 topic 失败:`);
        }
        // console.log(` try ${t} 次，获取随机内容失败:`);
        return null;
        
    } catch (error) {
        console.error(' error 获取随机 topic 失败:', error);
        return null;
    }
}

const testTweet = async() => {
    excelDataList = await getExcelDataList(file);
    if (!Array.isArray(excelDataList) || excelDataList.length === 0) {
        console.log(` 请配置 tweets.xlsx 内容表格 `);
        return false;
    }
    await sleep(3000);
    const contents = ( await getRamdomContent() ) + " " + ( await getRamdomAt() ) + "  " + ( await getRamdomTopic() ) + " " ;
    console.log(`推特完整内容：${contents}`);
}

const like = async (itemData) => {
    if(!itemData){
        console.log("请配置数据表格");
        return false;
    }
    
    const browser = await pBrowser.getBrowserByGlobalConfig(itemData.windowName);
    if(!browser){
        console.log(`can't get browser: ${itemData.windowName}`);
        return false;
    }

    await sleep(3000);
    const pages = await browser.pages();
    let page = await pages.find(mpage => mpage.url().includes("x.com")) || null;
    if(!page){
        page = await browser.newPage();
        await page.goto(xUrl, { waitUntil:"load" });
        await sleep(2000);
    }

    try {
        // Navigate to the tweet URL if provided
        if(itemData.tweetUrl) {
            await page.goto(itemData.tweetUrl, { waitUntil:"load" });
            await sleep(3000);
        } else {
            // If no specific tweet URL, go to home and like the first tweet
            await page.waitForSelector(`a[href="/home"]`, {timeout: 5000});
            await page.click(`a[href="/home"]`);
            await sleep(2000);
        }

        // 尝试多种 Like 按钮选择器
        const selectors = [
            '[data-testid="like"]',
            '[aria-label="Like"]',
            '[aria-label="喜欢"]',
            'div[data-testid="like"]',
            'div[aria-label="Like"]',
            'div[aria-label="喜欢"]',
            'svg[aria-label="Like"]',
            'svg[aria-label="喜欢"]'
        ];

        let clicked = false;
        for (const selector of selectors) {
            try {
                await page.waitForSelector(selector, { timeout: 4000 });
                await page.click(selector);
                clicked = true;
                break;
            } catch (e) {
                // 继续尝试下一个
            }
        }

        if (clicked) {
            await sleep(2000);
            return true;
        }

        console.log(`=== 没找到 Like 按钮 ===`);
        return false;

    } catch(err) {
        console.error('Like error:', err);
        return false;
    }
}

async function likeX(itemData) {
    try {
        console.log("== windowName = ", itemData.windowName);
        if (!itemData.tweetUrl) {
            console.log("没有 tweetUrl，无法点赞目标推文！");
            itemData.like = "NO_URL";
            await updateOneDataInExcel(excelDataList, itemData.id, itemData, file);
            return false;
        }
        const res = await like(itemData);
        itemData.like = res ? "TRUE" : "FALSE";
        await updateOneDataInExcel(excelDataList, itemData.id, itemData, file);
        await sleep(500);
        if (res) {
            console.log(` ${itemData.windowName}  like success `);
            return true;
        } else {
            console.log(` ${itemData.windowName}  like failed  `);
            return false;
        }
    } catch (err) {
        console.error(` ${itemData.windowName}  like err: ${err}  `);
        itemData.like = "ERROR";
        await updateOneDataInExcel(excelDataList, itemData.id, itemData, file);
        return false;
    }
}

async function getTweetUrls() {
    const tweetUrlList = await getExcelDataList(tweetUrlFile);
    // 只取 content 列且非空
    return tweetUrlList.map(item => item.content).filter(Boolean);
}

// ==================== 3LIAN 工具函数 ====================

// 日志计数器
let logCounter = 1;

/**
 * 封装的日志函数，自动添加精确时间戳和序号
 */
function logWithTime(message, category = "INFO", itemData = null) {
    const now = new Date();
    const timestamp = now.toISOString().replace('T', ' ').replace('Z', '');
    const milliseconds = now.getMilliseconds().toString().padStart(3, '0');
    const counterStr = `#${logCounter.toString().padStart(3, '0')}`;
    const identifier = itemData ? `[${itemData.windowName || itemData.xusername || 'Unknown'}]` : '';
    console.log(`[${timestamp}.${milliseconds}] [${category}] ${counterStr} ${identifier} ${message}`);
    logCounter++;
}

/**
 * 性能计时器，用于测量函数执行时间
 */
class PerformanceTimer {
    constructor() {
        this.timers = {};
    }
    
    start(label) {
        this.timers[label] = {
            startTime: Date.now(),
            endTime: null
        };
        logWithTime(`⏱️ 计时开始: ${label}`, "TIMER");
    }
    
    end(label) {
        if (!this.timers[label]) {
            logWithTime(`❌ 错误: 尝试结束未开始的计时器 ${label}`, "ERROR");
            return 0;
        }
        
        this.timers[label].endTime = Date.now();
        const duration = this.timers[label].endTime - this.timers[label].startTime;
        logWithTime(`⏱️ 计时结束: ${label} - 耗时: ${duration}ms (${(duration/1000).toFixed(2)}秒)`, "TIMER");
        return duration;
    }
}

/**
 * 获取随机回复内容（从reply.xlsx）
 */
async function getRandomReplyContent() {
    try {
        const replyFile = "./reply.xlsx";
        logWithTime(`📊 正在加载回复内容数据，文件路径: ${replyFile}`, "DATA");
        
        if (!fs.existsSync(replyFile)) {
            logWithTime(`⚠️ 回复内容Excel文件不存在: ${replyFile}，使用默认内容`, "WARNING");
            return "Great point! 👍";
        }
        
        const contentList = await getExcelDataList(replyFile);
        if (!contentList || contentList.length === 0) {
            logWithTime(`⚠️ 回复内容Excel文件为空，使用默认内容`, "WARNING");
            return "Great point! 👍";
        }
        
        logWithTime(`✅ 成功加载回复内容数据，共 ${contentList.length} 条记录`, "DATA");
        
        // 获取随机回复内容
        const validContents = contentList.filter(item => item.reply);
        let randomContent = "";
        if (validContents.length > 0) {
            randomContent = validContents[Math.floor(Math.random() * validContents.length)].reply;
            logWithTime(`📝 随机选择回复内容: ${randomContent}`, "RANDOM");
        } else {
            randomContent = "Great point! 👍";
            logWithTime(`⚠️ 未找到有效回复内容，使用默认: ${randomContent}`, "WARNING");
        }
        
        // 获取随机@用户
        const validAts = contentList.filter(item => item.at);
        let randomAt = "";
        if (validAts.length > 0) {
            randomAt = validAts[Math.floor(Math.random() * validAts.length)].at;
            logWithTime(`👤 随机选择@用户: ${randomAt}`, "RANDOM");
        }
        
        // 获取随机话题标签
        const validTopics = contentList.filter(item => item.topic);
        let randomTopic = "";
        if (validTopics.length > 0) {
            randomTopic = validTopics[Math.floor(Math.random() * validTopics.length)].topic;
            logWithTime(`🏷️ 随机选择话题标签: ${randomTopic}`, "RANDOM");
        }
        
        // 组合成完整的回复内容
        const fullContent = [randomContent, randomAt, randomTopic].filter(Boolean).join(" ");
        logWithTime(`✅ 组合完整内容: ${fullContent}`, "CONTENT");
        
        return fullContent || "Great point! 👍";
        
    } catch (err) {
        logWithTime(`❌ 获取随机回复内容失败: ${err.message}`, "ERROR");
        return "Great point! 👍";
    }
}

/**
 * 获取项目方推文数据
 */
async function getXiangmuData() {
    const xiangmuFile = "./data/xiangmu.xlsx";
    logWithTime(`📊 正在重新加载项目方推文数据，文件路径: ${xiangmuFile}`, "DATA");
    
    try {
        if (!fs.existsSync(xiangmuFile)) {
            logWithTime(`❌ 项目方Excel文件不存在: ${xiangmuFile}`, "ERROR");
            return [];
        }
        
        logWithTime(`🔄 强制重新读取Excel文件，获取最新数据...`, "REFRESH");
        const xiangmuDataList = await getExcelDataList(xiangmuFile);
        logWithTime(`✅ 成功重新加载项目方数据，共 ${xiangmuDataList.length} 条记录`, "DATA");
        
        if (xiangmuDataList.length == 0) {
            logWithTime(`⚠️ 项目方Excel文件为空或格式不正确: ${xiangmuFile}`, "WARNING");
            return [];
        }
        
        // 过滤出有URL的记录（不再检查3lian字段，因为要反复循环使用）
        const filteredArray = xiangmuDataList.filter(object => {
            if (!object.url) {
                logWithTime(`⚠️ 记录缺少url字段，跳过`, "WARNING");
                return false;
            }
            
            logWithTime(`🎯 记录 (${object.xm || 'Unknown'}) 符合处理条件，URL: ${object.url}`, "DATA");
            return true;
        });
        
        logWithTime(`🎯 符合条件的待处理推文数: ${filteredArray.length} 条`, "DATA");
        return filteredArray;
        
    } catch (error) {
        logWithTime(`❌ 加载项目方数据出错: ${error.message}`, "ERROR");
        return [];
    }
}

/**
 * 获取X账号数据 - 与x-login.js保持一致，不预先分配windowName
 */
async function getXAccountData() {
    const xAccountFile = "./data/x.xlsx";
    logWithTime(`📊 正在加载X账号数据，文件路径: ${xAccountFile}`, "DATA");
    
    try {
        const xAccountDataList = await getExcelDataList(xAccountFile);
        logWithTime(`✅ 成功加载X账号数据，共 ${xAccountDataList.length} 条记录`, "DATA");
        
        if (xAccountDataList.length == 0) {
            logWithTime(`⚠️ X账号Excel文件为空或格式不正确: ${xAccountFile}`, "WARNING");
            return [];
        }
        
        // 过滤出有用户名和密码，且3lian字段不为true的账号
        const filteredArray = xAccountDataList.filter(object => {
            if (!object.xusername || !object.xpassword) {
                logWithTime(`⚠️ 账号 ID:${object.id} 缺少用户名或密码，跳过`, "WARNING");
                return false;
            }
            
            // 检查3lian字段，如果为true则跳过（加强检查逻辑）
            const threeLianValue = object['3lian'];
            const isThreeLianTrue = (
                threeLianValue === true || 
                threeLianValue === 'true' || 
                threeLianValue === 'TRUE' || 
                threeLianValue === 'True' ||
                threeLianValue === 1 ||
                threeLianValue === '1' ||
                (typeof threeLianValue === 'string' && threeLianValue.toLowerCase().trim() === 'true')
            );
            
            if (isThreeLianTrue) {
                logWithTime(`⚠️ 账号 ID:${object.id} (${object.xusername}) 的3lian字段为true (值: ${threeLianValue})，跳过`, "SKIP");
                return false;
            }
            
            logWithTime(`🎯 账号 ID:${object.id} (${object.xusername}) 可用，3lian字段: ${threeLianValue || '空'}`, "DATA");
            return true;
        });
        
        // 不再预先分配windowName，在实际使用时动态分配
        logWithTime(`🎯 可用的X账号数: ${filteredArray.length} 个 (已过滤3lian=true的账号)`, "DATA");
        return filteredArray;
        
    } catch (error) {
        logWithTime(`❌ 加载X账号数据出错: ${error.message}`, "ERROR");
        return [];
    }
}

/**
 * 打开浏览器 - 与x-login.js保持一致
 */
async function openBrowser(itemData) {
    logWithTime(`🌐 正在打开浏览器窗口 ${itemData.windowName}...`, "BROWSER", itemData);
    
    try {
        console.log("windowName", itemData.windowName);
        const browser = await pBrowser.getBrowserByName(itemData.windowName);
        
        if (browser) {
            const browserObj = { [itemData.windowName]: browser };
            global.config.browsers.push(browserObj);
            logWithTime(`✅ 打开 ${itemData.windowName} 成功`, "BROWSER", itemData);
            return true;
        } else {
            logWithTime(`❌ 打开 ${itemData.windowName} 失败`, "ERROR", itemData);
            return false;
        }
    } catch (err) {
        logWithTime(`❌ 打开浏览器时发生错误: ${err.message}`, "ERROR", itemData);
        return false;
    }
}

/**
 * 关闭所有浏览器
 */
async function closeAllBrowser() {
    logWithTime(`🧹 开始关闭所有浏览器窗口...`, "CLEANUP");
    
    try {
        global.config.browsers.forEach(object => {
            for (let key in object) {
                const bro = object[key];
                bro.close();
                logWithTime(`✅ winName ${key} 关闭窗口成功`, "CLEANUP");
            }
        });
        global.config.browsers = [];
        logWithTime(`✅ 所有浏览器窗口关闭完成`, "CLEANUP");
    } catch (error) {
        logWithTime(`⚠️ 关闭浏览器时发生错误: ${error.message}`, "WARNING");
    }
}

/**
 * 获取账号对应的浏览器实例
 */
async function getBrowserForAccount(accountData) {
    try {
        if (!global.config || !global.config.browsers) {
            logWithTime(`❌ 全局浏览器配置未初始化`, "ERROR", accountData);
            return null;
        }
        
        const windowName = accountData.windowName;
        
        // 在已打开的浏览器中查找对应的窗口
        for (const browserObj of global.config.browsers) {
            if (browserObj[windowName]) {
                logWithTime(`✅ 找到窗口 ${windowName} 对应的浏览器实例`, "BROWSER", accountData);
                return browserObj[windowName];
            }
        }
        
        logWithTime(`❌ 未找到窗口 ${windowName} 对应的浏览器实例`, "ERROR", accountData);
        return null;
        
    } catch (error) {
        logWithTime(`❌ 获取浏览器实例失败: ${error.message}`, "ERROR", accountData);
        return null;
    }
}

/**
 * 更新X账号记录 - 只更新3lian字段
 */
async function updateXAccountRecord(accountData, success, accountDataList) {
    try {
        const xAccountFile = "./data/x.xlsx";
        logWithTime(`💾 开始更新X账号记录...`, "UPDATE", accountData);
        logWithTime(`📋 账号ID: ${accountData.id}, 用户名: ${accountData.xusername}`, "UPDATE", accountData);
        
        if (success === 1) {
            // 成功时设置3lian字段为1
            accountData['3lian'] = 1;
            logWithTime(`✅ 设置3lian字段为1（成功）`, "UPDATE", accountData);
        } else if (success === 0) {
            // 检测到cloudfire时设置3lian字段为0
            accountData['3lian'] = 0;
            logWithTime(`⚠️ 设置3lian字段为0（检测到cloudfire）`, "UPDATE", accountData);
        }
        
        await updateOneDataInExcel(accountDataList, accountData.id, accountData, xAccountFile);
        logWithTime(`✅ X账号记录更新完成`, "UPDATE", accountData);
        
    } catch (error) {
        logWithTime(`❌ 更新X账号记录失败: ${error.message}`, "ERROR", accountData);
    }
}

/**
 * 随机选择一条推文数据
 */
function getRandomTweet(tweetDataList) {
    if (!tweetDataList || tweetDataList.length === 0) {
        logWithTime(`❌ 推文数据列表为空，无法随机选择`, "ERROR");
        return null;
    }
    
    const randomIndex = Math.floor(Math.random() * tweetDataList.length);
    const selectedTweet = tweetDataList[randomIndex];
    
    logWithTime(`🎲 随机选择推文 #${randomIndex + 1}/${tweetDataList.length}`, "RANDOM");
    logWithTime(`📋 选中推文: 项目=${selectedTweet.xm || 'Unknown'}`, "RANDOM");
    logWithTime(`🔗 推文链接: ${selectedTweet.url}`, "RANDOM");
    
    return selectedTweet;
}

/**
 * 从推文URL提取用户名
 */
function extractUsernameFromTweetUrl(tweetUrl) {
    try {
        const urlMatch = tweetUrl.match(/(?:x\.com|twitter\.com)\/([^\/]+)\/status/);
        if (urlMatch && urlMatch[1]) {
            const username = urlMatch[1];
            logWithTime(`🔍 从URL提取用户名: ${username}`, "EXTRACT");
            return username;
        } else {
            logWithTime(`⚠️ 无法从URL提取用户名: ${tweetUrl}`, "WARNING");
            return null;
        }
    } catch (error) {
        logWithTime(`❌ 提取用户名失败: ${error.message}`, "ERROR");
        return null;
    }
}

/**
 * 记录已关注的用户到文件
 */
async function recordFollowedUser(accountData, username) {
    try {
        const { checkOrCreatePath } = require("../../utils/file.js");
        
        // 创建followed-txt目录（如果不存在）
        const followedDir = path.resolve(__dirname, './followed-txt');
        if (!fs.existsSync(followedDir)) {
            fs.mkdirSync(followedDir, { recursive: true });
            logWithTime(`📁 创建followed-txt目录: ${followedDir}`, "CREATE", accountData);
        }
        
        // 创建或获取关注记录文件路径
        const followedPath = path.resolve(__dirname, `./followed-txt/${accountData.windowName}.txt`);
        await checkOrCreatePath(followedPath);
        
        // 将用户名写入文件
        fs.writeFileSync(followedPath, `${username}\n`, { flag: 'a' });
        logWithTime(`✅ 成功记录关注用户: ${username} 到文件: ${followedPath}`, "RECORD", accountData);
        
        return true;
    } catch (error) {
        logWithTime(`❌ 记录关注用户失败: ${error.message}`, "ERROR", accountData);
        return false;
    }
}

/**
 * 检查当前登录的X账户用户名
 */
async function getCurrentLoggedInUsername(page, accountData) {
    try {
        logWithTime(`🔍 检查当前登录的账户用户名...`, "CHECK", accountData);
        
        if (!page || page.isClosed()) {
            logWithTime(`❌ 页面不可用或已关闭`, "ERROR", accountData);
            return null;
        }
        
        await page.goto('https://x.com/home', { 
            waitUntil: 'networkidle2', 
            timeout: 30000 
        });
        await sleep(3000);
        
        const currentUrl = page.url();
        if (currentUrl.includes('/login') || currentUrl.includes('/i/flow/login')) {
            logWithTime(`ℹ️ 当前在登录页面，无账户登录`, "INFO", accountData);
            return null;
        }
        
        let currentUsername = null;
        
        // 尝试点击用户头像按钮获取用户名
        try {
            const avatarSelectors = [
                'a[data-testid="AppTabBar_Profile_Link"]',
                '[data-testid="SideNav_AccountSwitcher_Button"]',
                'div[data-testid="SideNav_AccountSwitcher_Button"]',
                'button[data-testid="SideNav_AccountSwitcher_Button"]'
            ];
            
            let avatarButton = null;
            for (const selector of avatarSelectors) {
                try {
                    await page.waitForSelector(selector, { timeout: 2000 });
                    avatarButton = await page.$(selector);
                    if (avatarButton) {
                        logWithTime(`✅ 找到用户头像按钮: ${selector}`, "FOUND", accountData);
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (avatarButton) {
                await avatarButton.click();
                await sleep(2000);
                
                const allElements = await page.$$('span, div');
                for (const element of allElements) {
                    try {
                        const text = await page.evaluate(el => el.textContent, element);
                        if (text && text.startsWith('@') && text.length > 1 && text.length < 50) {
                            const username = text.substring(1).trim();
                            if (username && /^[a-zA-Z0-9_]+$/.test(username)) {
                                currentUsername = username;
                                logWithTime(`✅ 通过点击头像找到用户名: ${currentUsername}`, "AVATAR", accountData);
                                break;
                            }
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                await page.click('body');
                await sleep(500);
            }
        } catch (e) {
            logWithTime(`⚠️ 通过头像按钮获取用户名失败: ${e.message}`, "WARNING", accountData);
        }
        
        if (currentUsername) {
            logWithTime(`✅ 成功获取当前登录用户名: ${currentUsername}`, "SUCCESS", accountData);
            return currentUsername;
        } else {
            logWithTime(`❌ 无法获取当前登录用户名`, "ERROR", accountData);
            return null;
        }
        
    } catch (error) {
        logWithTime(`❌ 检查当前登录用户名失败: ${error.message}`, "ERROR", accountData);
        return null;
    }
}

/**
 * 退出X账号登录
 */
async function logoutX(page, accountData) {
    try {
        logWithTime(`🚪 开始退出当前X账号...`, "LOGOUT", accountData);
        
        if (!page || page.isClosed()) {
            logWithTime(`❌ 页面不可用或已关闭`, "ERROR", accountData);
            return false;
        }
        
        await page.goto('https://x.com/home', { 
            waitUntil: 'networkidle2', 
            timeout: 30000 
        });
        await sleep(2000);
        
        // 查找用户菜单按钮
        const menuSelectors = [
            '[data-testid="SideNav_AccountSwitcher_Button"]',
            '[aria-label="Account menu"]',
            '[data-testid="AppTabBar_More_Menu"]'
        ];
        
        let menuButton = null;
        for (const selector of menuSelectors) {
            try {
                await page.waitForSelector(selector, { timeout: 3000 });
                menuButton = await page.$(selector);
                if (menuButton) {
                    logWithTime(`✅ 找到菜单按钮: ${selector}`, "FOUND", accountData);
                    break;
                }
            } catch (e) {
                continue;
            }
        }
        
        if (menuButton) {
            await menuButton.click();
            await sleep(1000);
        }
        
        // 查找退出登录选项
        const logoutSelectors = [
            'a[href="/logout"]',
            '[data-testid="AccountSwitcher_Logout_Button"]'
        ];
        
        let logoutOption = null;
        for (const selector of logoutSelectors) {
            try {
                await page.waitForSelector(selector, { timeout: 2000 });
                logoutOption = await page.$(selector);
                if (logoutOption) {
                    logWithTime(`✅ 找到退出登录选项: ${selector}`, "FOUND", accountData);
                    break;
                }
            } catch (e) {
                continue;
            }
        }
        
        if (!logoutOption) {
            logWithTime(`⚠️ 未找到退出登录按钮，尝试直接访问退出URL`, "WARNING", accountData);
            await page.goto('https://x.com/logout', { 
                waitUntil: 'networkidle2', 
                timeout: 30000 
            });
            await sleep(2000);
            
            try {
                const confirmButton = await page.$('button[data-testid="confirmationSheetConfirm"]');
                if (confirmButton) {
                    await confirmButton.click();
                    await sleep(2000);
                    logWithTime(`✅ 确认退出登录`, "SUCCESS", accountData);
                }
            } catch (e) {
                logWithTime(`⚠️ 无法找到确认退出按钮`, "WARNING", accountData);
            }
        } else {
            await logoutOption.click();
            await sleep(2000);
            
            try {
                const confirmButton = await page.$('button[data-testid="confirmationSheetConfirm"]');
                if (confirmButton) {
                    await confirmButton.click();
                    await sleep(2000);
                    logWithTime(`✅ 确认退出登录`, "SUCCESS", accountData);
                }
            } catch (e) {
                // 可能不需要确认，直接退出了
            }
        }
        
        await sleep(3000);
        const currentUrl = page.url();
        if (currentUrl.includes('/login') || currentUrl.includes('/i/flow/login') || !currentUrl.includes('/home')) {
            logWithTime(`✅ 成功退出登录，当前URL: ${currentUrl}`, "SUCCESS", accountData);
            return true;
        } else {
            logWithTime(`⚠️ 退出登录可能未成功，当前URL: ${currentUrl}`, "WARNING", accountData);
            return false;
        }
        
    } catch (error) {
        logWithTime(`❌ 退出登录失败: ${error.message}`, "ERROR", accountData);
        return false;
    }
}

/**
 * 执行推文四连操作（关注、点赞、转发、评论）
 */
async function performThreeLianActions(tweetUrl, accountData, page, combinedContent = null) {
    logWithTime(`🎯 开始执行四连操作...`, "ACTION", accountData);
    
    let successCount = 0;
    const results = {
        follow: false,
        like: false,
        retweet: false,
        reply: false
    };
    
    try {
        if (!page || page.isClosed()) {
            logWithTime(`❌ 页面不可用或已关闭`, "ERROR", accountData);
            return { success: false, error: "页面不可用", results, successCount };
        }
        
        // 访问推文页面
        logWithTime(`🔗 准备访问推文页面: ${tweetUrl}`, "NAVIGATE", accountData);
        await page.goto(tweetUrl, { 
            waitUntil: 'networkidle2', 
            timeout: 30000 
        });
        logWithTime(`✅ 推文页面加载完成`, "NAVIGATE", accountData);
        
        await sleep(3000);
        
        // 检查页面是否正常加载
        const currentUrl = page.url();
        logWithTime(`📍 当前页面URL: ${currentUrl}`, "INFO", accountData);
        
        if (currentUrl.includes('suspended') || currentUrl.includes('account/access')) {
            logWithTime(`⚠️ 账户可能被暂停或需要验证`, "WARNING", accountData);
            return { success: false, error: "账户被暂停或需要验证", results, successCount };
        }
        
        // 1. 关注操作
        logWithTime(`👥 开始执行关注操作...`, "ACTION", accountData);
        const targetUsername = extractUsernameFromTweetUrl(tweetUrl);
        
        try {
            const followSelectors = [
                '[data-testid="follow"]',
                'button[aria-label*="Follow"]',
                'button:has-text("Follow")',
                'button:has-text("关注")'
            ];
            
            let followButton = null;
            for (const selector of followSelectors) {
                try {
                    await page.waitForSelector(selector, { timeout: 2000 });
                    followButton = await page.$(selector);
                    if (followButton) {
                        logWithTime(`✅ 找到关注按钮: ${selector}`, "FOUND", accountData);
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (followButton) {
                const buttonText = await page.evaluate(btn => btn.textContent, followButton);
                if (buttonText.includes('Following') || buttonText.includes('已关注')) {
                    logWithTime(`ℹ️ 已经关注该用户，跳过关注操作`, "INFO", accountData);
                    results.follow = true;
                    successCount++;
                    
                    if (targetUsername) {
                        await recordFollowedUser(accountData, targetUsername);
                    }
                } else {
                    await followButton.click();
                    await sleep(2000);
                    logWithTime(`✅ 关注操作完成`, "SUCCESS", accountData);
                    results.follow = true;
                    successCount++;
                    
                    if (targetUsername) {
                        await recordFollowedUser(accountData, targetUsername);
                        logWithTime(`📝 已记录关注用户: ${targetUsername}`, "RECORD", accountData);
                    }
                }
            } else {
                logWithTime(`⚠️ 未找到关注按钮，可能已经关注或页面结构变化`, "WARNING", accountData);
                results.follow = true;
                successCount++;
                
                if (targetUsername) {
                    await recordFollowedUser(accountData, targetUsername);
                }
            }
            
            // 检查是否已经完成了3个操作
            if (successCount >= 3) {
                logWithTime(`✅ 已完成至少3个操作，提前结束四连操作`, "SUCCESS", accountData);
                return { 
                    success: true, 
                    successCount, 
                    results,
                    error: null
                };
            }
            
        } catch (error) {
            logWithTime(`❌ 关注操作失败: ${error.message}`, "ERROR", accountData);
        }
        
        // 2. 点赞操作
        logWithTime(`❤️ 开始执行点赞操作...`, "ACTION", accountData);
        try {
            const likeSelectors = [
                '[data-testid="like"]',
                'button[aria-label*="Like"]',
                'button[aria-label*="点赞"]'
            ];
            
            let likeButton = null;
            for (const selector of likeSelectors) {
                try {
                    await page.waitForSelector(selector, { timeout: 2000 });
                    likeButton = await page.$(selector);
                    if (likeButton) {
                        logWithTime(`✅ 找到点赞按钮: ${selector}`, "FOUND", accountData);
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (likeButton) {
                await likeButton.click();
                await sleep(2000);
                logWithTime(`✅ 点赞操作完成`, "SUCCESS", accountData);
                results.like = true;
                successCount++;
            } else {
                logWithTime(`❌ 未找到点赞按钮`, "ERROR", accountData);
            }
            
            // 检查是否已经完成了3个操作
            if (successCount >= 3) {
                logWithTime(`✅ 已完成至少3个操作，提前结束四连操作`, "SUCCESS", accountData);
                return { 
                    success: true, 
                    successCount, 
                    results,
                    error: null
                };
            }
            
        } catch (error) {
            logWithTime(`❌ 点赞操作失败: ${error.message}`, "ERROR", accountData);
        }
        
        // 3. 转发操作（优先引用转发，失败则直接转发）
        logWithTime(`🔄 开始执行转发操作...`, "ACTION", accountData);
        try {
            await sleep(2000);
            
            const retweetSelectors = [
                '[data-testid="retweet"]',
                'button[aria-label*="Repost"]',
                'button[aria-label*="转发"]'
            ];
            
            let retweetButton = null;
            for (const selector of retweetSelectors) {
                try {
                    await page.waitForSelector(selector, { timeout: 3000 });
                    retweetButton = await page.$(selector);
                    if (retweetButton) {
                        const isVisible = await page.evaluate(btn => {
                            const rect = btn.getBoundingClientRect();
                            return rect.width > 0 && rect.height > 0 && !btn.disabled;
                        }, retweetButton);
                        
                        if (isVisible) {
                            logWithTime(`✅ 找到可用的转发按钮: ${selector}`, "FOUND", accountData);
                            break;
                        } else {
                            retweetButton = null;
                        }
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (retweetButton) {
                await retweetButton.click();
                await sleep(2000);
                
                // 检查转发菜单
                const menuItems = await page.$$('[role="menuitem"]');
                
                // 优先尝试引用转发
                let quoteOption = null;
                for (const item of menuItems) {
                    try {
                        const text = await page.evaluate(el => el.textContent.trim(), item);
                        if (text === "Quote" || text === "引用推文") {
                            quoteOption = item;
                            logWithTime(`✅ 找到Quote选项`, "FOUND", accountData);
                            break;
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                if (quoteOption) {
                    await quoteOption.click();
                    await sleep(2000);
                    
                    // 在引用中添加内容 - 使用传入的combinedContent或随机内容
                    let replyContent;
                    if (combinedContent) {
                        replyContent = combinedContent;
                        logWithTime(`📝 使用组合内容作为引用: ${replyContent.substring(0, 50)}${replyContent.length > 50 ? '...' : ''}`, "CONTENT", accountData);
                    } else {
                        replyContent = await getRandomReplyContent();
                        logWithTime(`📝 使用随机内容作为引用: ${replyContent}`, "CONTENT", accountData);
                    }
                    
                    const textareaSelectors = [
                        '[data-testid="tweetTextarea_0"]',
                        'div[contenteditable="true"]',
                        'div[role="textbox"]'
                    ];
                    
                    let textarea = null;
                    for (const selector of textareaSelectors) {
                        try {
                            await page.waitForSelector(selector, { timeout: 3000 });
                            textarea = await page.$(selector);
                            if (textarea) {
                                logWithTime(`✅ 找到文本输入框: ${selector}`, "FOUND", accountData);
                                break;
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    if (textarea) {
                        await textarea.click();
                        await sleep(500);
                        await textarea.type(replyContent, { delay: 50 });
                        await sleep(1000);
                        
                        await textarea.focus();
                        await sleep(500);
                        
                        // 发送引用推文
                        try {
                            await page.keyboard.down('Control');
                            await page.keyboard.press('Enter');
                            await page.keyboard.up('Control');
                            await sleep(3000);
                            logWithTime(`✅ 使用快捷键发送引用推文成功`, "SUCCESS", accountData);
                            results.retweet = true;
                            successCount++;
                        } catch (shortcutError) {
                            // 备用方案：点击发送按钮
                            const postSelectors = [
                                'div[data-testid="tweetButton"]',
                                '[data-testid="tweetButton"]',
                                'button[data-testid="tweetButton"]'
                            ];
                            
                            let postButton = null;
                            for (const selector of postSelectors) {
                                try {
                                    await page.waitForSelector(selector, { timeout: 3000 });
                                    postButton = await page.$(selector);
                                    if (postButton) {
                                        const isEnabled = await page.evaluate(btn => !btn.disabled, postButton);
                                        if (isEnabled) {
                                            logWithTime(`✅ 找到可用的发送按钮: ${selector}`, "FOUND", accountData);
                                            break;
                                        }
                                    }
                                } catch (e) {
                                    continue;
                                }
                            }
                            
                            if (postButton) {
                                await postButton.click();
                                await sleep(3000);
                                logWithTime(`✅ 点击Post按钮发送成功`, "SUCCESS", accountData);
                                results.retweet = true;
                                successCount++;
                            }
                        }
                    }
                } else {
                    // 直接转发
                    let directRetweetOption = null;
                    for (const item of menuItems) {
                        try {
                            const text = await page.evaluate(el => el.textContent.trim(), item);
                            if (text === "Repost" || text === "转发") {
                                directRetweetOption = item;
                                logWithTime(`✅ 找到Repost选项`, "FOUND", accountData);
                                break;
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    if (directRetweetOption) {
                        await directRetweetOption.click();
                        await sleep(2000);
                        logWithTime(`✅ 直接转发操作完成`, "SUCCESS", accountData);
                        results.retweet = true;
                        successCount++;
                    }
                }
            }
            
            // 检查是否已经完成了3个操作
            if (successCount >= 3) {
                logWithTime(`✅ 已完成至少3个操作，提前结束四连操作`, "SUCCESS", accountData);
                return { 
                    success: true, 
                    successCount, 
                    results,
                    error: null
                };
            }
            
        } catch (error) {
            logWithTime(`❌ 转发操作失败: ${error.message}`, "ERROR", accountData);
        }
        
        // 4. 评论操作（在转发后进行）
        if (results.retweet) {
            logWithTime(`💬 开始执行评论操作...`, "ACTION", accountData);
            try {
                const replySelectors = [
                    '[data-testid="reply"]',
                    'button[aria-label*="Reply"]',
                    'button[aria-label*="回复"]'
                ];
                
                let replyButton = null;
                for (const selector of replySelectors) {
                    try {
                        await page.waitForSelector(selector, { timeout: 3000 });
                        replyButton = await page.$(selector);
                        if (replyButton) {
                            const isVisible = await page.evaluate(btn => {
                                const rect = btn.getBoundingClientRect();
                                return rect.width > 0 && rect.height > 0 && !btn.disabled;
                            }, replyButton);
                            
                            if (isVisible) {
                                logWithTime(`✅ 找到可用的回复按钮: ${selector}`, "FOUND", accountData);
                                break;
                            } else {
                                replyButton = null;
                            }
                        }
                    } catch (e) {
                        continue;
                    }
                }
                
                if (replyButton) {
                    await replyButton.click();
                    await sleep(2000);
                    
                    // 使用传入的combinedContent或者获取随机回复内容
                    let replyContent;
                    if (combinedContent) {
                        replyContent = combinedContent;
                        logWithTime(`📝 使用组合内容作为回复: ${replyContent.substring(0, 50)}${replyContent.length > 50 ? '...' : ''}`, "CONTENT", accountData);
                    } else {
                        replyContent = await getRandomReplyContent();
                        logWithTime(`📝 使用随机内容作为回复: ${replyContent}`, "CONTENT", accountData);
                    }
                    
                    const replyTextareaSelectors = [
                        '[data-testid="tweetTextarea_0"]',
                        'div[contenteditable="true"]',
                        'div[role="textbox"]'
                    ];
                    
                    let replyTextarea = null;
                    for (const selector of replyTextareaSelectors) {
                        try {
                            await page.waitForSelector(selector, { timeout: 3000 });
                            replyTextarea = await page.$(selector);
                            if (replyTextarea) {
                                logWithTime(`✅ 找到回复文本输入框: ${selector}`, "FOUND", accountData);
                                break;
                            }
                        } catch (e) {
                            continue;
                        }
                    }
                    
                    if (replyTextarea) {
                        await replyTextarea.click();
                        await sleep(500);
                        await replyTextarea.type(replyContent, { delay: 50 });
                        await sleep(1000);
                        
                        await replyTextarea.focus();
                        await sleep(500);
                        
                        // 发送回复
                        try {
                            await page.keyboard.down('Control');
                            await page.keyboard.press('Enter');
                            await page.keyboard.up('Control');
                            await sleep(3000);
                            logWithTime(`✅ 使用快捷键发送回复成功`, "SUCCESS", accountData);
                            results.reply = true;
                            successCount++;
                        } catch (shortcutError) {
                            // 备用方案：点击发送按钮
                            const replyPostSelectors = [
                                'div[data-testid="tweetButton"]',
                                '[data-testid="tweetButton"]',
                                'button[data-testid="tweetButton"]'
                            ];
                            
                            let replyPostButton = null;
                            for (const selector of replyPostSelectors) {
                                try {
                                    await page.waitForSelector(selector, { timeout: 3000 });
                                    replyPostButton = await page.$(selector);
                                    if (replyPostButton) {
                                        const isEnabled = await page.evaluate(btn => !btn.disabled, replyPostButton);
                                        if (isEnabled) {
                                            logWithTime(`✅ 找到可用的回复发送按钮: ${selector}`, "FOUND", accountData);
                                            break;
                                        }
                                    }
                                } catch (e) {
                                    continue;
                                }
                            }
                            
                            if (replyPostButton) {
                                await replyPostButton.click();
                                await sleep(3000);
                                logWithTime(`✅ 点击Reply按钮发送成功`, "SUCCESS", accountData);
                                results.reply = true;
                                successCount++;
                            }
                        }
                    }
                }
            } catch (error) {
                logWithTime(`❌ 评论操作失败: ${error.message}`, "ERROR", accountData);
            }
        }
        
        // 判断四连操作是否成功（4个操作中至少成功3个）
        const isSuccess = successCount >= 3;
        logWithTime(`📊 四连操作完成，成功 ${successCount}/4 个操作`, "SUMMARY", accountData);
        logWithTime(`📋 操作结果: 关注=${results.follow}, 点赞=${results.like}, 转发=${results.retweet}, 评论=${results.reply}`, "SUMMARY", accountData);
        
        return { 
            success: isSuccess, 
            successCount, 
            results,
            error: isSuccess ? null : "部分操作失败" 
        };
        
    } catch (error) {
        logWithTime(`❌ 四连操作过程中发生错误: ${error.message}`, "ERROR", accountData);
        return { success: false, successCount, results, error: error.message };
    }
}

// 导出新增的函数
module.exports = {
    loginAgain,
    login,
    follow,
    tweet,
    like,
    likeX,
    getTweetUrls,
    testTweet,
    getRamdomContent,
    getRamdomAt,
    getRamdomTopic,
    // 新增的3lian工具函数
    logWithTime,
    PerformanceTimer,
    getRandomReplyContent,
    getXiangmuData,
    getXAccountData,
    openBrowser,
    closeAllBrowser,
    getBrowserForAccount,
    updateXAccountRecord,
    getRandomTweet,
    extractUsernameFromTweetUrl,
    recordFollowedUser,
    getCurrentLoggedInUsername,
    logoutX,
    performThreeLianActions
};