const axios = require('axios');
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
let file = './file.xlsx';
async function main() {
    let sheets = await readXlsxFile(file, { getSheets: true });
    for(let i=0; i<sheets.length; i++){
        await downloadInSheet(sheets[i].name, i+1);
    }

    //await downloadImage();
}

async function downloadInSheet(sheetName, index){
    let data = await readXlsxFile(file, { sheet: index });
    for(let i=0; i<data.length; i++){
        let row = data[i];
        if(!row[0]) continue;
        let filename = `${sheetName} - ${row[0]}`;
        let url = row[10] || '';
        if(url.match(/drive\.google\.com/gi)){
            url = url.split('/');
            let id = url[url.length-2];
            if(!id) continue;
            url = `https://drive.google.com/u/0/uc?id=${id}&export=download`;
            //console.log(url)
        }
        if(!url) continue;
        console.log(filename, url)
        if(url.match(/drive\.google\.com/gi) || url.match(/http:/gi)){
            try{
                await downloadImage(url, filename);
            }
            catch(err){
                console.log(err.message);
            }
        }
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
        writer.on('finish', resolve)
        writer.on('error', reject)
    })
}

main();