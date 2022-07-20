const { Transaction, AccessListEIP2930Transaction } = require("@ethereumjs/tx");
import {Account, Address, BN, bufferToHex, isValidAddress, toBuffer} from 'ethereumjs-util'
import { recordTxStatus } from './api';
const whiteList = require("./whitelist.json")
const EventEmitter = require('events');
const axios = require("axios");
const config = require("./config.json")
const fs = require('fs')
export let node = {
    ip: 'localhost',
    port: 9001
}

let verbose = config.verbose
let gotArchiver = false
let nodeList: any[] = []
let nextIndex = 0

export async function updateNodeList() {
    if (config.askLocalHostForArchiver === true) {
        if (gotArchiver === false) {
            gotArchiver = true
            //TODO query a localhost (or other) node or a valid archiver IP
        }
    }

    const res = await axios.get(`http://${config.archiverIpInfo.externalIp}:${config.archiverIpInfo.externalPort}/nodelist`)
    const nodes = res.data.nodeList
    if (nodes.length > 0) {
        nodeList = [...nodes]
        if (verbose) console.log('Nodelist is updated')
    }
}

export async function waitRandomSecond() {
    let second = Math.floor(Math.random() * 5) + 1
    if (verbose) console.log(`Waiting ${second} second`)
    await sleep(second * 1000)
}

export async function requestWithRetry(method: string, url: string, data: any = {}) {
    let retry = 0 
    let maxRetry = 5 //set this to 0 with for load testing rpc server
    let success = false
    while (!success && retry <= maxRetry) {
        retry++
        try {
            // if (true) console.log(`Running request with retry: ${url} count: ${retry}`)
            const res = await axios({
                method,
                url,
                data
            });
            if (res.status === 200 && !res.data.error) {
                success = true
                return res
            }
        } catch (e: any) {
            console.log('Error: requestWithRetry', e.message)
        }
        
        if(retry <= maxRetry){
            if (verbose) console.log('Node is busy...will try again in a few seconds')
            await waitRandomSecond()            
        } else {
            if (verbose) console.log('Node is busy...out of retries')
        }
    }
    return { data: null }
}

export function getTransactionObj(tx: any): any {
    if (!tx.raw) throw Error('No raw tx found.')
    let transactionObj
    const serializedInput = toBuffer(tx.raw)
    try {
        transactionObj = Transaction.fromRlpSerializedTx(serializedInput)
        if (verbose) console.log('Legacy tx parsed:', transactionObj)
    } catch (e) {
        if (verbose) console.log('Unable to get legacy transaction obj', e)
    }
    if (!transactionObj) {
        try {
            transactionObj = AccessListEIP2930Transaction.fromRlpSerializedTx(serializedInput)
            if (verbose) console.log('EIP2930 tx parsed:', transactionObj)
        } catch (e) {
            console.log('Unable to get EIP2930 transaction obj', e)
        }
    }

    if (transactionObj) {
        return transactionObj
    } else throw Error('tx obj fail')
}

export function intStringToHex(str: string) {
    return '0x' + parseInt(str, 10).toString(16)
}
export function getBaseUrl() {
    setConsensorNode()
    return `http://${node.ip}:${node.port}`
}

export function changeNode(ip: string, port: number) {
    node.ip = ip
    node.port = port
    if (verbose) console.log(`RPC server subscribes to ${ip}:${port}`)
}

function rotateConsensorNode() {
    let consensor: any = getNextConsensorNode()//getRandomConsensorNode()
    if (consensor) {
        let nodeIp = consensor.ip
        //Sometimes the external IPs returned will be local IPs.  This happens with pm2 hosting multpile nodes on one server.
        //config.useConfigNodeIp will override the local IPs with the config node external IP when rotating nodes
        if (config.useConfigNodeIp === true) {
            nodeIp = config.nodeIpInfo.externalIp
        }
        changeNode(nodeIp, consensor.port)
    }
}

// export function apiStatCollector(methodName: any, args: string[]) {
//     let now = Math.round(Date.now() / 1000)
//     if (perfTracker[methodName]) {
//         perfTracker[methodName].push(true)
//     } else {
//         perfTracker[methodName] = [true]
//     }
// }

// this is the main function to be called every RPC request
export function setConsensorNode() {
    if (config.dynamicConsensorNode) {
        rotateConsensorNode()
    } else {
        changeNode(config.nodeIpInfo.externalIp, config.nodeIpInfo.externalPort)
    }
}

export function getRandomConsensorNode() {
    if (nodeList.length > 0) {
        let randomIndex = Math.floor(Math.random() * nodeList.length)
        return nodeList[randomIndex]
    }
}

/**
 * Round robin selection of next consensor index.
 * @returns 
 */
export function getNextConsensorNode() {
    if (nodeList.length > 0) {
        nextIndex++
        if(nextIndex >= nodeList.length){
            nextIndex = 0
        }
        return nodeList[nextIndex]
    }
}

export function sleep(ms: number) {
    return new Promise(resolve => {
        setTimeout(() => {
            resolve(true)
        }, ms)
    })
}

export async function getAccount(addressStr: any) {
    try {
        const url = getBaseUrl();
        if (verbose) console.log(`${url}/account/${addressStr}`)
        // let res = await axios.get(`${getBaseUrl()}/account/${addressStr}`)
        let res = await requestWithRetry('get', `${url}/account/${addressStr}`)
        return res.data.account
    } catch (e) {
        // console.log('getAccount error', e)
    }
}

export class RequestersList {
  heavyRequests: Map<string, number[]>
  heavyAddresses: Map<string, number[]>
  abusedToAddresses: any
  bannedIps: any[]
  requestTracker: any
  allRequestTracker: any
  totalTxTracker: any

  constructor(blackList: string[] = []) {
    this.heavyRequests = new Map()
    this.heavyAddresses = new Map()
    this.abusedToAddresses = {}
    this.requestTracker = {}
    this.allRequestTracker = {}
    this.totalTxTracker = {}
    this.bannedIps = blackList.map((ip: string) => {
      return {ip, timestamp: Date.now()}
    })
    let self = this
    setInterval(() => {
      self.clearOldIps()
    }, 60 * 1000)
    setInterval(() => {
      self.logMostFrequentIps()
    }, 5 * 60 * 1000)
  }

  addToBlacklist(ip: string) {
    this.bannedIps.push({ip, timestamp: Date.now()})
    fs.readFile('blacklist.json', function (err: any, currentDataStr: string) {
      const ipList = JSON.parse(currentDataStr)
      if (ipList.indexOf(ip) >= 0) return
      let newIpList = [...ipList, ip]
      console.log(`Added ip ${ip} to banned list`)
      fs.writeFileSync('blacklist.json', JSON.stringify(newIpList))
    })
  }

  clearOldIps() {
    const now = Date.now()
    const oneMinute = 60 * 1000
    for (let [ip, reqHistory] of this.heavyRequests) {
      console.log(`In last 60s, IP ${ip} made ${reqHistory.length} heavy requests`)
      // for (let j = 0; j < reqHistory.length; j++) {
      //   if (j > 0) {
      //     console.log('time delta between reqs', reqHistory[j] - reqHistory[j - 1], 'ms')
      //   }
      // }
    }
    for (let [ip, reqHistory] of this.heavyRequests) {
      let i = 0
      for (; i < reqHistory.length; i++) {
        if (now - reqHistory[i] < oneMinute) break // we can stop looping the record array here
      }
      if (i > 0) reqHistory.splice(0, i - 1) // oldest item is at index 0
      //console.log('reqHistory after clearing heavy request history', reqHistory.length)
    }

    for (let [address, reqHistory] of this.heavyAddresses) {
      let i = 0
      for (; i < reqHistory.length; i++) {
        if (now - reqHistory[i] < oneMinute) break // we can stop looping the record array here
      }
      if (i > 0) reqHistory.splice(0, i - 1) // oldest item is at index 0
      //console.log('reqHistory after clearing heavy request history', reqHistory.length)
    }

    // unban the ip after 1 hour
    this.bannedIps = this.bannedIps.filter((record: any) => {
      if (now - record.timestamp >= 60 * 60 * 1000) return false
      else return true
    })
  }

  logMostFrequentIps() {
    // log and clean successful requests
    let records = Object.values(this.requestTracker)
    records = records.sort((a: any, b: any) => b.count - a.count)
    if (config.verbose) console.log('Most frequent successful IPs:', records)
    this.requestTracker = {}

    // log and clean all requests
    let allRecords = Object.values(this.allRequestTracker)
    allRecords = allRecords.sort((a: any, b: any) => b.count - a.count)
    if (config.verbose) console.log('Most frequent all IPs (rejected + successful):', allRecords)
    this.allRequestTracker = {}

    // log total injected tx by ip
    let txRecords = Object.values(this.totalTxTracker)
    txRecords = txRecords.sort((a: any, b: any) => b.count - a.count)
    for (let i = 0; i < txRecords.length; i++) {
      let txRecord: any = txRecords[i]
      if (txRecord.count >= 20) {
        if (whiteList.indexOf(txRecord.ip) === -1) {
          console.log('ban this ip due to continuously heavy requests')
          // this.addToBlacklist(txRecord.ip)
        }
      }
    }
    console.log('Total num of txs injected by IPs', txRecords)
    this.totalTxTracker = {}

    // log abused contract addresses
    let mostAbusedSorted: any[] = Object.values(this.abusedToAddresses).sort((a: any, b: any) => b.count - a.count)
    for (let abusedData of mostAbusedSorted) {
      console.log(`Contract address: ${abusedData.to}. Count: ${abusedData.count}`)
      console.log(`Most frequent caller addresses:`)
      let sortedCallers: any[] = Object.values(abusedData.from).sort((a: any, b: any) => b.count - a.count)
      for (let caller of sortedCallers) {
        console.log(`    ${caller.from}, count: ${caller.count}`)
        let sortedIps: any[] = Object.values(caller.ips).sort((a: any, b: any) => b.count - a.count)
        for (let ip of sortedIps) {
          console.log(`             ${ip.ip}, count: ${ip.count}`)
        }
      }
      console.log('------------------------------------------------------------')
    }
  }

  addHeavyRequest(ip: string) {
    if (this.requestTracker[ip]) {
      this.requestTracker[ip].count += 1
    } else {
      this.requestTracker[ip] = {ip, count: 1}
    }
    if (this.totalTxTracker[ip]) {
      this.totalTxTracker[ip].count += 1
    } else {
      this.totalTxTracker[ip] = {ip, count: 1}
    }
    if (this.heavyRequests.get(ip)) {
      let reqHistory = this.heavyRequests.get(ip)
      if (reqHistory) reqHistory.push(Date.now())
    } else {
      this.heavyRequests.set(ip, [Date.now()])
    }
  }

  addHeavyAddress(address: string) {
    if (this.heavyAddresses.get(address)) {
      let reqHistory = this.heavyAddresses.get(address)
      if (reqHistory) reqHistory.push(Date.now())
    } else {
      this.heavyAddresses.set(address, [Date.now()])
    }
  }

  addAbusedAddress(toAddress: string, fromAddress: string, ip: string) {
    if (this.abusedToAddresses[toAddress]) {
      this.abusedToAddresses[toAddress].count += 1
      let fromData = this.abusedToAddresses[toAddress].from[fromAddress]
      if (fromData) {
        fromData.count += 1
        fromData.from = fromAddress
        if (fromData.ips[ip]) {
          fromData.ips[ip].count += 1
        } else {
          fromData.ips[ip] = {ip, count: 1}
        }
      } else {
        let newFromData: any = {
          count: 1,
          from: fromAddress,
          ips: {}
        }
        newFromData.ips[ip] = {
          count: 1,
          ip,
        }
        this.abusedToAddresses[toAddress].from[fromAddress] = newFromData
      }
    } else {
      this.abusedToAddresses[toAddress] = {
        to: toAddress,
        count: 1,
        from: {},
      }
      let newFromData: any = {
        count: 1,
        from: fromAddress,
        ips: {}
      }
      newFromData.ips[ip] = {
        count: 1,
        ip,
      }
      this.abusedToAddresses[toAddress].from[fromAddress] = newFromData
    }
  }

  addAllRequest(ip: string) {
    if (this.allRequestTracker[ip]) {
      this.allRequestTracker[ip].count += 1
    } else {
      this.allRequestTracker[ip] = {ip, count: 1}
    }
  }

  isIpBanned(ip: string) {
    let bannedIpList = this.bannedIps.map(data => data.ip)
    if (bannedIpList.indexOf(ip) >= 0) return true
    else return false
  }

  isQueryType(reqType: string, reqParams: any[]) {
    try {
      let heavyTypes = ['eth_sendRawTransaction', 'eth_sendTransaction']
      if (heavyTypes.indexOf(reqType) >= 0) return false
      // if (reqType === 'eth_call' && reqParams[0].data.indexOf('0x70a08231') === -1) {
      //   if(config.verbose) console.log('Not a balance query eth_call. Considered as heavy.')
      //   return false
      // }
      return true
    } catch (e) {
      return true
    }
  }

  isRequestOkay(ip: string, reqType: string, reqParams: any[]): boolean {
    const now = Date.now()
    const oneMinute = 60 * 1000

    if (whiteList.indexOf(ip) >= 0) return true

    if (this.isIpBanned(ip)) {
      console.log(`This ip ${ip} is banned.`, reqType, reqParams)
      return false
    }

    if (this.isQueryType(reqType, reqParams)) {
      return true
    }

    // record this heavy request before checking
    this.addHeavyRequest(ip)
    let heavyReqHistory = this.heavyRequests.get(ip)

    if (heavyReqHistory && heavyReqHistory.length >= 61) {
      if (now - heavyReqHistory[heavyReqHistory.length - 61] < oneMinute) {
        if (verbose) console.log(`Ban this ip ${ip} due to continuously sending more than 60 reqs in 60s`)
        // this.addToBlacklist(ip)
        return false
      }
    }

    let transaction
    try {
      if (reqType === 'eth_sendRawTransaction') transaction = getTransactionObj({raw: reqParams[0]})
    } catch (e) {

    }

    if (heavyReqHistory && heavyReqHistory.length >= 10) {
      if (now - heavyReqHistory[heavyReqHistory.length - 10] < oneMinute) {
        if (verbose) console.log(`Your last heavy req is less than 60s ago`, `total requests: ${heavyReqHistory.length}, `, Math.round((now - heavyReqHistory[heavyReqHistory.length - 10]) / 1000), 'seconds')
        if (transaction) {
          console.log('tx rejected', bufferToHex(transaction.hash()))
          if (config.recordTxStatus) recordTxStatus({
            txHash: bufferToHex(transaction.hash()),
            injected: false,
            accepted: false,
            reason: 'Rejected by JSON RPC rate limiting'
          })
        }
        return false
      }
    }

    if (reqType === 'eth_sendRawTransaction') {
      try {
        let readableTx = {
          from: transaction.getSenderAddress().toString(),
          to: transaction.to ? transaction.to.toString() : '',
          value: transaction.value.toString(),
          data: bufferToHex(transaction.data),
          hash: bufferToHex(transaction.hash())
        }
        if (readableTx.from) this.addHeavyAddress(readableTx.from)
        if (readableTx.to && readableTx.to !== readableTx.from) this.addHeavyAddress(readableTx.to)

        let fromAddressHistory = this.heavyAddresses.get(readableTx.from)
        if (fromAddressHistory && fromAddressHistory.length >= 10) {
          if (now - fromAddressHistory[fromAddressHistory.length - 10] < oneMinute) {
            if (verbose) console.log(`Your last req FROM this address ${readableTx.from} is less than 60s ago`, `total requests: ${fromAddressHistory.length}, `, Math.round((now - fromAddressHistory[fromAddressHistory.length - 10]) / 1000), 'seconds')
            if (config.recordTxStatus) recordTxStatus({
              txHash: bufferToHex(transaction.hash()),
              injected: false,
              accepted: false,
              reason: 'Rejected by JSON RPC rate limiting'
            })
            return false
          }
        }

        let toAddressHistory = this.heavyAddresses.get(readableTx.to)
        if (toAddressHistory && toAddressHistory.length >= 10) {
          if (now - toAddressHistory[toAddressHistory.length - 10] < oneMinute) {
            this.addAbusedAddress(readableTx.to, readableTx.from, ip)
            if (verbose) console.log(`Your last req TO this address ${readableTx.to} is less than 60s ago`, `total requests: ${toAddressHistory.length}, `, Math.round((now - toAddressHistory[toAddressHistory.length - 10]) / 1000), 'seconds')
            if (config.recordTxStatus) recordTxStatus({
              txHash: bufferToHex(transaction.hash()),
              injected: false,
              accepted: false,
              reason: 'Rejected by JSON RPC rate limiting'
            })
            return false
          }
        }
      } catch (e) {
        console.log('Error while get tx obj', e)
      }
    }
    if (heavyReqHistory && config.verbose) console.log(`We allow ip ${ip}`)
    return true
  }
}
