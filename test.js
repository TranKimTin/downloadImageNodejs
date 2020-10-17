const axios = require('axios');
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const puppeteer = require('puppeteer');
let file = './file.xlsx';
let success = 0, fail = 0;

let browser, page;
async function main() {
    browser = await puppeteer.launch({headless: false});
    page = await browser.newPage();
    createFolder('./image');
    let sheets = await readXlsxFile(file, { getSheets: true });
    for (let i = 0; i < sheets.length; i++) {
        await downloadInSheet(sheets[i].name, i + 1);
    }
    sleep(5000);
    moveFile();
    console.log("finish");
    await browser.close();
}

async function moveFile(){
    fs.readdirSync("./image").forEach(folder => {
        if(!folder.includes(".")){
            fs.readdirSync(`./image/${folder}`).forEach(file => {
                fs.renameSync(`./image/${folder}/${file}`, `./image/${folder}.png`, function(err) {
                    if (err) console.log('ERROR: ' + err);
                });
            });
            fs.rmdirSync(`./image/${folder}`, { recursive: true });
        }
    });
}

function createFolder(path) {
    if (!fs.existsSync(path)) {
        fs.mkdirSync(path);
    }
}

async function downloadInSheet(sheetName, index) {
    createFolder(`./image/${sheetName}`);
    let data = await readXlsxFile(file, { sheet: index });
    for (let i = 0; i < data.length; i++) {
        let row = data[i];
        if (!row[0]) continue;
        let filename = row[0] || 'name';
        let url = row[10] || '';
        if (url.match(/drive\.google\.com/gi)) {
            url = url.split('/');
            let id = url[url.length - 2];
            if (!id) continue;
            url = `http://drive.google.com/u/0/uc?id=${id}&export=download`;
            //console.log(url)
        }
        if (!url) continue;
        // url = url.replace("https://", "http://");
        console.log(filename, url)
        if (url.match(/drive\.google\.com/gi) || url.match(/http:/gi) || url.match(/https:/gi)) {
            await download1(url, filename, 0);
        }
    }
}
async function download1(url, filename){
    try {
        await sleep(1000);
        if(url.match(/drive\.google\.com/gi)){
            await downloadFromDriver(url, filename);
        }
        else{
            await downloadImage(url, filename);
        }
        console.log("SUCCESS", ++success, "FAIL", fail);
    }
    catch (err) {
        console.log(err.message);
        console.log("SUCCESS", success, "FAIL", ++fail);
    }
}
async function downloadImage(url, filename) {
    //'https://drive.google.com/u/0/uc?id=&export=download';
    //const url = 'https://drive.google.com/u/0/uc?id=1NOr9y2N_TStOsZgxWeoqVRW7cUcHmypU&export=download';
    const writer = fs.createWriteStream(`./image/${filename}.png`)
    const response = await axios({
        url,
        method: 'GET',
        responseType: 'stream'
    });
    response.data.pipe(writer)

    return new Promise((resolve, reject) => {
        writer.on('finish', resolve())
        writer.on('error', reject())
    })
}

async function downloadFromDriver(url, filename){
    try{
        let downloadPath = path.resolve(__dirname, `image/${filename}`);
        await page._client.send('Page.setDownloadBehavior', {behavior: 'allow', downloadPath: downloadPath});
        await go(url2);
    }
    catch(err){
        console.log(err.message);
    }
}

async function sleep(time) {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            resolve();
        }, time)
    })
}

main();