const express = require('express');
const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;
let scrapingStatus = { status: 'ready', currentPage: 0 };

// Middleware для роботи з JSON та формами
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static('public'));
app.set('view engine', 'ejs');

// Рендер головної сторінки
app.get('/', (req, res) => {
  res.render('index');
});

// Обробник статусу
app.get('/status', (req, res) => {
  res.json(scrapingStatus);
});

// Обробник форми парсингу
app.post('/scrape', async (req, res) => {
  const { url, pages } = req.body;
  scrapingStatus = { status: 'working', currentPage: 0 };

  try {
    // Функція для парсингу декількох сторінок
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.goto(url, { waitUntil: 'networkidle2' });

    let allData = [];
    let currentPage = 1;

    // Визначаємо, скільки сторінок парсити
    while (pages === 'all' || currentPage <= parseInt(pages)) {
      scrapingStatus.currentPage = currentPage; // Оновлюємо поточну сторінку
      console.log(`Scraping page ${currentPage}...`);

      // Парсинг поточної сторінки
      const pageData = await scrapePage(page);
      allData = allData.concat(pageData);

      const nextPageBtn = await page.$('.paginate__btn.next');
      if (!nextPageBtn || (pages !== 'all' && currentPage === parseInt(pages))) break; // Якщо кнопки для наступної сторінки немає або досягли ліміту сторінок

      // Перехід на наступну сторінку та очікування
      await nextPageBtn.click();
      await new Promise(resolve => setTimeout(resolve, 3000)); // Чекаємо завантаження нової сторінки

      currentPage++;
    }

    await browser.close();
    await saveDataToExcel(allData);

    scrapingStatus = { status: 'done', currentPage: currentPage - 1 }; // Оновлюємо статус
    res.redirect('/');
  } catch (error) {
    console.error('Error:', error.message);
    scrapingStatus = { status: 'ready', currentPage: 0 }; // Скидаємо статус
    res.redirect('/');
  }
});

// Обробник завантаження Excel-файлу
app.get('/download', (req, res) => {
  const file = path.join(__dirname, 'data.xlsx');
  res.download(file);
});

// Функція для парсингу однієї сторінки
async function scrapePage(page) {
  return await page.evaluate(() => {
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
}

// Функція для збереження даних в Excel
async function saveDataToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');

  worksheet.columns = [
    { header: 'Title', key: 'title', width: 30 },
    { header: 'Price', key: 'price', width: 20 },
    { header: 'Description', key: 'description', width: 50 },
    { header: 'ID', key: 'id', width: 30 },
    { header: 'Completed', key: 'completed', width: 25 }
  ];

  data.forEach((row, index) => {
    // console.log(`Adding row ${index + 1}:`, row);
    worksheet.addRow(row);
  });

  await workbook.xlsx.writeFile('data.xlsx');
  console.log('Data has been saved to data.xlsx');
}

// Запуск серверу
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
