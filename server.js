const express = require('express');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));

let isScraping = false;
let fileGenerated = false;

app.get('/', (req, res) => {
  res.render('index', { isScraping, fileGenerated });
});

app.post('/scrape', async (req, res) => {
  const url = req.body.url;
  isScraping = true;
  fileGenerated = false;

  try {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });

    const data = await page.evaluate(() => {
      const results = [];
      let currentItem = {};

      document.querySelectorAll('div.search-result-card__col').forEach((element) => {
        const title = element.querySelector('.item-title__title')?.innerText.trim() || '';
        const description = element.querySelector('.search-result-card__description')?.innerText.trim() || '';
        const idElement = element.querySelector('p.search-result-card__description');
        const id = idElement ? idElement.innerText.trim().replace('ID: ', '') : '';
        const price = element.querySelector('.app-price__amount')?.innerText.trim() || '';

        const statusElement = element.querySelector('.search-result-card__label span');
        const completed = statusElement ? statusElement.innerText.trim() : '';

        if (title && description) {
          if (currentItem.title) {
            results.push(currentItem);
          }
          currentItem = { title, description, id, completed, price };
        } else {
          if (price) {
            currentItem.price = price;
          }
          if (completed && currentItem.title) {
            currentItem.completed = completed;
          }
        }
      });

      if (currentItem.title) {
        results.push(currentItem);
      }

      return results;
    });

    await browser.close();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    worksheet.columns = [
      { header: 'Title', key: 'title', width: 30 },
      { header: 'Price', key: 'price', width: 20 },
      { header: 'Description', key: 'description', width: 50 },
      { header: 'ID', key: 'id', width: 30 },
      { header: 'Completed', key: 'completed', width: 25 }
    ];

    data.forEach((row) => {
      worksheet.addRow(row);
    });

    await workbook.xlsx.writeFile('public/data.xlsx');
    fileGenerated = true;
    isScraping = false;
  } catch (error) {
    console.error('Error:', error.message);
    isScraping = false;
  }

  res.redirect('/');
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
