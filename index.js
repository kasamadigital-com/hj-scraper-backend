const express = require('express');
const fs = require('fs');
const cors = require('cors');
const xlsx = require('xlsx');
const { Builder } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
require('chromedriver');
const path = require('path');

const app = express();
const port = 3000;
app.use(cors());
app.use(express.json());

// Timeout function to handle long-running operations
const withTimeout = (promise, ms) => {
  let timeoutId;
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(`Timeout after ${ms} ms`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timeoutId));
};

app.post('/excel', (req, res) => {
  let json = req.body.json;
  let keyword = req.body.keywords;
  let instanceNumber = req.body.instanceNumber;

  // Define the header of words you're searching for
  const headers = keyword.split('\n');
  const headers2 = keyword.split('\n');
  headers2.unshift('Website'); 
  
  // Initialize worksheet data with the headers
  const worksheetData = [headers2];

  // Loop through the data and construct each row
  json.forEach(entry => {
    const row = [entry.message];

    // For each word in the headers, except the first (which is 'Website'), check if it's found
    headers.forEach(word => {
      if (!entry.results || entry.results.length === 0) {
        row.push('');  // Add 'No Results' if entry has no results
      } else if (entry.results.includes(word)) {
        row.push('X');  // Add 'X' if the word is found
      } else {
        row.push('');  // Leave empty if the word is not found
      }
    });

    // Add the row to the worksheet data
    worksheetData.push(row);
  });

  // Create a new workbook and add the worksheet
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);

  // Append the worksheet to the workbook
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Findings');

  // Write the workbook to a file
  xlsx.writeFile(workbook, `./temp/findings_${instanceNumber}.xlsx`);

  const filePath = path.join(__dirname, 'temp', `findings_${instanceNumber}.xlsx`);

  fs.access(filePath, fs.constants.F_OK, (err) => {
    console.log(filePath)

    if (err) {
        console.error('File not found');
        return res.status(404).send('File not found');
    }

    // Set the headers for the file download
    res.setHeader('Content-Disposition', `attachment; filename="findings_${instanceNumber}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    // Stream the file to the client
    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);
});

})

// POST route to check the website content
app.post('/check', async (req, res) => {
  let KEYWORD = req.body.keyword;
  let URL = req.body.url;
  let INSTANCENUMBER = req.body.instanceNumber;

  const keywordArray = KEYWORD.split('\n');

  try {
    const results = await createDriverInstance(keywordArray, URL, INSTANCENUMBER);  // Run 5 instances of Selenium

    res.json({
      message: URL,
      instanceNumber: INSTANCENUMBER,
      results: results,
    });
  } catch (error) {
    console.error(error);
    res.status(500).send('Error running Selenium instances');
  }
});

async function createDriverInstance(keywordArray, url, instanceNumber) {
  const options = new chrome.Options();
  options.addArguments('--headless','--disable-gpu', '--no-sandbox', '--disable-dev-shm-usage', '--log-level=3', '--disable-extensions', '--disable-background-timer-throttling');

  let driver = await new Builder()
    .forBrowser('chrome')
    .setChromeOptions(options)
    .build();

  try {
    console.log(`Instance ${instanceNumber} running`);

    await withTimeout(driver.get(`https://${url}`), 20000);

    await withTimeout(driver.wait(async () => {
      const readyState = await driver.executeScript('return document.readyState');
      return readyState === 'complete';
    }, 10000), 20000);

    const pageSource = await driver.getPageSource();
    const wordsToSearch = keywordArray;

    const foundStrings = wordsToSearch.filter(substring => pageSource.includes(substring));
    
    return foundStrings;
  }catch (error) {
    console.error('Error:', error.message);
  } finally {
    await driver.quit();
    console.log(`Instance ${instanceNumber} closed`);
  }
}

app.get('/', (req, res) => res.send({ status: 'ok' }));
app.get('/test', (req, res) => res.send({ status: 'test ok' }));

app.listen(port, () => console.log(`Server running on http://localhost:${port}`));