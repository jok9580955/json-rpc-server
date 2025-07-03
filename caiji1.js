/**
 * 项目方X最近推文采集脚本
 * 功能: 采集项目方X账户的最近一条推文内容和URL
 * 作者: Auto Biter
 * 创建时间: 2024
 */

const pBrowser = require("../../utils/pBrowser.js");
const { sleep, delay } = require("../../utils/time.js");
const path = require("path");
const fs = require("fs");
const { getExcelDataList, updateOneDataInExcel } = require("../../utils/ExcelDataUtils");

// 配置文件路径
const excelFile = "D:\\200\\000000000000000000\\auto_biter\\BrowserScript\\x\\data\\xiangmu.xlsx";

// 全局变量
var excelDataList = [];
var globalBrowser = null;
var globalPage = null;

/**
 * 封装的日志函数，自动添加精确时间戳和分类标签
 */
function logWithTime(message, category = "INFO") {
    const now = new Date();
    const timestamp = now.toISOString().replace('T', ' ').replace('Z', '');
    const milliseconds = now.getMilliseconds().toString().padStart(3, '0');
    console.log(`[${timestamp}.${milliseconds}] [${category}] ${message}`);
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

// 创建全局性能计时器实例
const performanceTimer = new PerformanceTimer();

/**
 * 初始化全局配置
 */
const initGlobalConfig = async () => {
    if (!global.config) {
        logWithTime('🔧 初始化全局配置...', "INIT");
        global.config = {};
        global.config.browsers = [];
        global.config.userList = await pBrowser.getAllUserList();
        logWithTime('✅ 全局配置初始化完成', "INIT");
    }
}

/**
 * 获取Excel数据
 */
async function getExcelData() {
    performanceTimer.start('getExcelData');
    logWithTime(`📊 正在加载Excel数据，文件路径: ${excelFile}`, "DATA");
    
    try {
        // 检查Excel文件是否存在
        if (!fs.existsSync(excelFile)) {
            logWithTime(`❌ Excel文件不存在: ${excelFile}`, "ERROR");
            return [];
        }
        
        // 加载Excel数据
        excelDataList = await getExcelDataList(excelFile);
        logWithTime(`✅ 成功加载Excel数据，共 ${excelDataList.length} 条记录`, "DATA");
        
        if (excelDataList.length == 0) {
            logWithTime(`⚠️ Excel文件为空或格式不正确: ${excelFile}`, "WARNING");
            return [];
        }
        
        // 为每条记录添加唯一ID（如果没有的话）
        excelDataList.forEach((item, index) => {
            if (!item.id) {
                item.id = index + 1; // 从1开始的行号作为ID
                logWithTime(`🔧 为记录 ${index + 1} 添加ID: ${item.id}`, "DATA");
            }
        });
        
        // 检查数据结构
        if (excelDataList.length > 0) {
            const firstItem = excelDataList[0];
            logWithTime(`📋 数据示例 - 第一条记录: ID=${firstItem.id}, xuser=${firstItem.xuser}, caiji=${firstItem.caiji}`, "DATA");
            
            // 检查必要字段
            if (!firstItem.xuser) {
                logWithTime(`⚠️ 警告: 数据可能缺少xuser字段，请检查Excel文件格式`, "WARNING");
            }
            
            // 显示所有字段名
            const fieldNames = Object.keys(firstItem);
            logWithTime(`📋 Excel字段列表: ${fieldNames.join(', ')}`, "DATA");
        }
        
        // 过滤出符合条件的记录：有xuser字段，且caiji字段不是true
        const filteredArray = excelDataList.filter(object => {
            // 必须有xuser字段
            if (!object.xuser) {
                logWithTime(`⚠️ 记录ID ${object.id} 缺少xuser字段，跳过`, "WARNING");
                return false;
            }
            
            // 如果caiji字段是true，则跳过
            if (object.caiji === true || object.caiji === 'true' || object.caiji === 1 || object.caiji === 'TRUE') {
                logWithTime(`✅ 记录ID ${object.id} (${object.xuser}) 的caiji字段已经是true，跳过处理`, "SKIP");
                return false;
            }
            
            // 其他情况都处理
            logWithTime(`🎯 记录ID ${object.id} (${object.xuser}) 符合处理条件，caiji字段值: ${object.caiji}`, "DATA");
            return true;
        });
        
        logWithTime(`🎯 符合条件的待处理记录数: ${filteredArray.length} 条`, "DATA");
        performanceTimer.end('getExcelData');
        return filteredArray;
        
    } catch (error) {
        logWithTime(`❌ 加载Excel数据出错: ${error.message}`, "ERROR");
        logWithTime(`📋 错误详情: ${error.stack}`, "ERROR");
        performanceTimer.end('getExcelData');
        return [];
    }
}

/**
 * 初始化浏览器（只打开一次）
 */
async function initBrowser() {
    performanceTimer.start('initBrowser');
    logWithTime(`🌐 正在初始化浏览器...`, "BROWSER");
    
    try {
        await initGlobalConfig();
        
        // 使用第一个可用的窗口
        const windowName = 1; // 固定使用窗口1
        logWithTime(`🔍 尝试获取浏览器窗口: ${windowName}`, "BROWSER");
        
        let browser = await pBrowser.getBrowserByGlobalConfig(windowName);
        
        if (!browser) {
            logWithTime(`🔍 未找到现有浏览器实例，尝试创建新的浏览器窗口 ${windowName}`, "BROWSER");
            browser = await pBrowser.getBrowserByName(windowName);
        }
        
        if (browser) {
            globalBrowser = browser;
            let browserObj = {};
            browserObj[windowName] = browser;
            global.config.browsers.push(browserObj);
            
            // 创建页面
            globalPage = await browser.newPage();
            
            // 设置用户代理，避免被检测
            await globalPage.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36');
            
            // 等待浏览器完全稳定
            await sleep(3000);
            
            logWithTime(`✅ 浏览器窗口 ${windowName} 初始化成功`, "BROWSER");
            performanceTimer.end('initBrowser');
            return true;
        } else {
            logWithTime(`❌ 浏览器窗口 ${windowName} 初始化失败`, "ERROR");
            performanceTimer.end('initBrowser');
            return false;
        }
    } catch (err) {
        logWithTime(`❌ 初始化浏览器时发生错误: ${err.message}`, "ERROR");
        performanceTimer.end('initBrowser');
        return false;
    }
}

/**
 * 采集项目方X的最近一条推文
 */
async function collectLatestTweet(itemData) {
    performanceTimer.start(`collectTweet_${itemData.xuser}`);
    logWithTime(`🐦 开始采集用户 ${itemData.xuser} 的最近推文...`, "COLLECT");
    
    try {
        if (!globalPage) {
            logWithTime(`❌ 全局页面未初始化`, "ERROR");
            return { success: false, error: "全局页面未初始化" };
        }
        
        // 清理用户名（移除@符号）
        const cleanUsername = itemData.xuser.replace('@', '');
        
        // 构建用户主页URL
        const userUrl = `https://x.com/${cleanUsername}`;
        logWithTime(`🔗 准备访问用户主页: ${userUrl}`, "NAVIGATE");
        
        // 访问用户主页
        performanceTimer.start(`pageLoad_${cleanUsername}`);
        await globalPage.goto(userUrl, { 
            waitUntil: 'networkidle2', 
            timeout: 60000 
        });
        performanceTimer.end(`pageLoad_${cleanUsername}`);
        logWithTime(`✅ 页面加载完成: ${userUrl}`, "NAVIGATE");
        
        await sleep(6000); // 等待页面完全加载
        logWithTime(`⏳ 页面稳定等待完成`, "WAIT");
        
        // 检查页面是否正常加载
        const currentUrl = globalPage.url();
        logWithTime(`📍 当前页面URL: ${currentUrl}`, "INFO");
        
        if (currentUrl.includes('suspended') || currentUrl.includes('account/access')) {
            logWithTime(`⚠️ 用户账户可能被暂停或需要验证: ${itemData.xuser}`, "WARNING");
            return { success: false, error: "账户被暂停或需要验证" };
        }
        
        // 采集最近一条推文前再等待一下
        await sleep(3000);
        logWithTime(`🔍 开始查找最近的推文...`, "COLLECT");
        
        // 采集最近一条推文
        const tweetData = await globalPage.evaluate(() => {
            // 查找推文容器
            const tweetSelectors = [
                'article[data-testid="tweet"]',
                '[data-testid="tweet"]',
                'article[role="article"]'
            ];
            
            let tweetElement = null;
            for (const selector of tweetSelectors) {
                const elements = document.querySelectorAll(selector);
                if (elements.length > 0) {
                    // 找到第一条推文（最新的）
                    tweetElement = elements[0];
                    break;
                }
            }
            
            if (!tweetElement) {
                console.log('未找到推文元素');
                return null;
            }
            
            // 提取推文内容
            let tweetText = '';
            const textSelectors = [
                '[data-testid="tweetText"]',
                '[lang] span',
                'div[dir="auto"] span'
            ];
            
            for (const selector of textSelectors) {
                const textElements = tweetElement.querySelectorAll(selector);
                if (textElements.length > 0) {
                    tweetText = Array.from(textElements)
                        .map(el => el.textContent)
                        .join(' ')
                        .trim();
                    if (tweetText) break;
                }
            }
            
            // 提取推文URL
            let tweetUrl = '';
            const linkElements = tweetElement.querySelectorAll('a[href*="/status/"]');
            if (linkElements.length > 0) {
                const href = linkElements[0].getAttribute('href');
                if (href) {
                    tweetUrl = href.startsWith('http') ? href : `https://x.com${href}`;
                }
            }
            
            // 提取推文时间
            let tweetTime = '';
            const timeElements = tweetElement.querySelectorAll('time');
            if (timeElements.length > 0) {
                tweetTime = timeElements[0].getAttribute('datetime') || timeElements[0].textContent;
            }
            
            console.log(`找到推文 - 内容长度: ${tweetText.length}, URL: ${tweetUrl}`);
            
            return {
                text: tweetText,
                url: tweetUrl,
                time: tweetTime
            };
        });
        
        if (tweetData && (tweetData.text || tweetData.url)) {
            logWithTime(`✅ 成功采集到推文数据`, "SUCCESS");
            logWithTime(`📝 推文内容长度: ${tweetData.text ? tweetData.text.length : 0} 字符`, "INFO");
            logWithTime(`🔗 推文URL: ${tweetData.url || '未获取到'}`, "INFO");
            logWithTime(`⏰ 推文时间: ${tweetData.time || '未获取到'}`, "INFO");
            
            performanceTimer.end(`collectTweet_${itemData.xuser}`);
            
            return {
                success: true,
                data: {
                    content: tweetData.text || '',
                    url: tweetData.url || '',
                    time: tweetData.time || ''
                }
            };
        } else {
            logWithTime(`⚠️ 未找到有效的推文数据`, "WARNING");
            performanceTimer.end(`collectTweet_${itemData.xuser}`);
            return { success: false, error: "未找到推文数据" };
        }
        
    } catch (error) {
        logWithTime(`❌ 采集推文时发生错误: ${error.message}`, "ERROR");
        performanceTimer.end(`collectTweet_${itemData.xuser}`);
        return { success: false, error: error.message };
    }
}

/**
 * 更新Excel数据
 */
async function updateExcelRecord(itemData, success, tweetData = null) {
    performanceTimer.start(`updateExcel_${itemData.excelRowNumber}`);
    logWithTime(`💾 开始更新Excel记录 行号: ${itemData.excelRowNumber}`, "UPDATE");
    
    try {
        // 检查行号是否存在
        if (itemData.rowIndex === undefined) {
            logWithTime(`❌ 记录缺少rowIndex字段，无法更新`, "ERROR");
            performanceTimer.end(`updateExcel_${itemData.excelRowNumber}`);
            return false;
        }
        
        // 更新数据对象
        const updatedData = { ...itemData };
        
        if (success && tweetData) {
            updatedData.caiji = true; // 标记采集成功
            updatedData.url = tweetData.url; // 保存推文URL
            updatedData.content = tweetData.content; // 保存推文内容
            updatedData.time = tweetData.time; // 保存推文时间
            logWithTime(`✅ 准备更新成功状态 - URL: ${tweetData.url}`, "UPDATE");
        } else {
            // 采集失败时也要标记，避免重复处理
            updatedData.caiji = false; // 标记采集失败
            logWithTime(`⚠️ 采集失败，caiji字段设置为false`, "UPDATE");
        }
        
        // 检查excelDataList中是否存在对应rowIndex的记录
        if (itemData.rowIndex >= excelDataList.length || itemData.rowIndex < 0) {
            logWithTime(`❌ rowIndex ${itemData.rowIndex} 超出范围，excelDataList长度: ${excelDataList.length}`, "ERROR");
            performanceTimer.end(`updateExcel_${itemData.excelRowNumber}`);
            return false;
        }
        
        logWithTime(`🔍 找到记录索引: ${itemData.rowIndex}，准备更新`, "UPDATE");
        
        // 直接更新内存中的数据列表
        excelDataList[itemData.rowIndex] = updatedData;
        logWithTime(`✅ 内存数据列表更新完成`, "UPDATE");
        
        // 更新Excel文件 - 使用rowIndex作为假的ID
        await updateOneDataInExcel(excelDataList, itemData.rowIndex, updatedData, excelFile);
        logWithTime(`✅ Excel记录更新完成 行号: ${itemData.excelRowNumber}`, "UPDATE");
        
        performanceTimer.end(`updateExcel_${itemData.excelRowNumber}`);
        return true;
        
    } catch (error) {
        logWithTime(`❌ 更新Excel记录失败 行号: ${itemData.excelRowNumber}, 错误: ${error.message}`, "ERROR");
        logWithTime(`📋 错误堆栈: ${error.stack}`, "ERROR");
        performanceTimer.end(`updateExcel_${itemData.excelRowNumber}`);
        return false;
    }
}

/**
 * 处理单个项目方账户
 */
async function processAccount(itemData) {
    const accountTimer = `processAccount_${itemData.xuser}`;
    performanceTimer.start(accountTimer);
    logWithTime(`🚀 开始处理账户: ${itemData.xuser} (ID: ${itemData.id})`, "PROCESS");
    
    try {
        // 采集推文
        const result = await collectLatestTweet(itemData);
        
        if (result.success) {
            logWithTime(`✅ 账户 ${itemData.xuser} 推文采集成功`, "SUCCESS");
            await updateExcelRecord(itemData, true, result.data);
            performanceTimer.end(accountTimer);
            return true;
        } else {
            logWithTime(`❌ 账户 ${itemData.xuser} 推文采集失败: ${result.error}`, "ERROR");
            await updateExcelRecord(itemData, false);
            performanceTimer.end(accountTimer);
            return false;
        }
        
    } catch (error) {
        logWithTime(`❌ 处理账户 ${itemData.xuser} 时发生未预期错误: ${error.message}`, "ERROR");
        await updateExcelRecord(itemData, false);
        performanceTimer.end(accountTimer);
        return false;
    }
}

/**
 * 清理资源
 */
async function cleanup() {
    logWithTime(`🧹 开始清理资源...`, "CLEANUP");
    
    try {
        if (globalPage) {
            await globalPage.close();
            logWithTime(`✅ 全局页面已关闭`, "CLEANUP");
        }
        
        // 注意：不关闭浏览器，保持浏览器实例运行
        logWithTime(`✅ 资源清理完成`, "CLEANUP");
    } catch (error) {
        logWithTime(`⚠️ 清理资源时发生错误: ${error.message}`, "WARNING");
    }
}

/**
 * 主函数
 */
async function main() {
    const mainTimer = 'main_execution';
    performanceTimer.start(mainTimer);
    logWithTime(`🎬 ========== 项目方X推文采集脚本启动 ==========`, "MAIN");
    logWithTime(`📅 启动时间: ${new Date().toLocaleString()}`, "MAIN");
    
    // 初始等待，让系统完全准备好
    logWithTime(`⏳ 系统初始化等待...`, "MAIN");
    await sleep(2000);
    
    try {
        // 获取Excel数据
        const dataList = await getExcelData();
        if (dataList.length === 0) {
            logWithTime(`⚠️ 没有找到需要处理的数据，脚本结束`, "MAIN");
            performanceTimer.end(mainTimer);
            return;
        }
        
        // 初始化浏览器（只初始化一次）
        const browserInitSuccess = await initBrowser();
        if (!browserInitSuccess) {
            logWithTime(`❌ 浏览器初始化失败，脚本结束`, "MAIN");
            performanceTimer.end(mainTimer);
            return;
        }
        
        logWithTime(`📊 总共需要处理 ${dataList.length} 个账户`, "MAIN");
        
        let successCount = 0;
        let failCount = 0;
        
        // 逐个处理账户
        for (let i = 0; i < dataList.length; i++) {
            const itemData = dataList[i];
            const currentIndex = i + 1;
            
            logWithTime(`\n📋 ========== 处理进度: ${currentIndex}/${dataList.length} ==========`, "PROGRESS");
            logWithTime(`👤 当前账户: ${itemData.xuser} (ID: ${itemData.id})`, "PROGRESS");
            
            const success = await processAccount(itemData);
            
            if (success) {
                successCount++;
                logWithTime(`✅ 账户处理成功 (${successCount}/${currentIndex})`, "PROGRESS");
            } else {
                failCount++;
                logWithTime(`❌ 账户处理失败 (${failCount}/${currentIndex})`, "PROGRESS");
            }
            
            // 添加延迟，避免请求过于频繁
            if (i < dataList.length - 1) {
                logWithTime(`⏳ 等待 6 秒后处理下一个账户...`, "PROGRESS");
                await sleep(6000);
            }
        }
        
        // 输出最终统计结果
        logWithTime(`\n🎉 ========== 采集任务完成 ==========`, "MAIN");
        logWithTime(`📊 总处理数量: ${dataList.length}`, "MAIN");
        logWithTime(`✅ 成功数量: ${successCount}`, "MAIN");
        logWithTime(`❌ 失败数量: ${failCount}`, "MAIN");
        logWithTime(`📈 成功率: ${((successCount / dataList.length) * 100).toFixed(2)}%`, "MAIN");
        
        performanceTimer.end(mainTimer);
        logWithTime(`🎉 脚本执行完成，程序退出`, "MAIN");
        
    } catch (error) {
        logWithTime(`💥 主函数执行过程中发生严重错误: ${error.message}`, "FATAL");
        logWithTime(`📋 错误堆栈: ${error.stack}`, "FATAL");
        performanceTimer.end(mainTimer);
    } finally {
        // 清理资源
        await cleanup();
    }
}

// 启动主函数
main().catch(error => {
    console.error(`💥 程序启动失败: ${error.message}`);
    console.error(`📋 错误堆栈: ${error.stack}`);
    process.exit(1);
});