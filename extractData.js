const enigma = require('enigma.js');
const schema = require('enigma.js/schemas/12.20.0.json');
const WebSocket = require('ws');
const fs = require('fs');
require('dotenv').config();

/* 
For this demo all sensitive information has been stored in a .env file.
The contents of this .env file are used/read by 'require('dotenv').config()', see row #5.
To keep this script working create a file and rename it to '.env'
Write the following:

APPID=[APPID_OF_APP]
URL=[SERVER ADDRESS]
TABLEOBJECTID=[OBJECT ID OF THE TABLE]

In this QS Desktop Demo I've used the following:

APPID=C:\\Users\\XYZ\\Documents\\Qlik\\Sense\\Apps\\Apex Global Operations - DRT Final.qvf
URL=ws://localhost:9076/app/engineData
TABLEOBJECTID=jKFVSU

*/

const config = {
  schema,
  appId: process.env.APPID, //replace with your appId, as this was tested on QS Desktop it currently has a direct path.
  url: process.env.URL, // replace with your server url
  createSocket: url => new WebSocket(url),
};

(async () => {
  try {
    const session = enigma.create(config);
    const qix = await session.open();
    const app = await qix.openDoc(config.appId);
    console.log('Opened App');
    // console.log(app);

    const tableObjectId = process.env.TABLEOBJECTID; // replace with your Table Object Id
    const tableObject = await app.getObject(tableObjectId);

    console.log('opened Table');

    const layout = await tableObject.getLayout();
    const qWidth = layout.qHyperCube.qSize.qcx;
    const qHeight = layout.qHyperCube.qSize.qcy;
    const qDimensionInfo = layout.qHyperCube.qDimensionInfo;
    const qMeasureInfo = layout.qHyperCube.qMeasureInfo;

    console.log(`Table dimensions: ${qWidth} columns x ${qHeight} rows`);

    const dataPages = await tableObject.getHyperCubeData('/qHyperCubeDef', [
      {
        qTop: 0,
        qLeft: 0,
        qWidth,
        qHeight,
      },
    ]);

    const rows = dataPages[0].qMatrix;

    // console.log('----------------------TABLEOBJECT----------------------------------')
    // console.log(tableObject)
    // console.log('----------------------ROWS----------------------------------')
    // console.log(rows)
   
    console.log('Writing table data to file')
    const csvFileName = `output_${new Date().toISOString().slice(0, 16).replace(/[-T:]/g, '')}.csv`;
    const csvFileContents = [
      [...qDimensionInfo.map(dim => dim.qFallbackTitle), ...qMeasureInfo.map(measure => measure.qFallbackTitle)].join(';')
    ].concat(rows.map(row => row.map(cell => cell.qText).join(';'))).join('\n');
    fs.writeFileSync(csvFileName, csvFileContents);
    console.log(`Output saved to file: ${csvFileName}`);

    await session.close();
  } catch (error) {
    console.error('Error:', error);
  }
})();
