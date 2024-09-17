const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
  try {
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();

    // URL сайта, который вы хотите спарсить
    const url = 'https://prozorro.gov.ua/uk/search/tender?cpv=15500000-3&page=500';
    await page.goto(url, { waitUntil: 'networkidle2' });

    // Извлечение данных
    const data = await page.evaluate(() => {
      const results = [];
      let currentItem = {};

      document.querySelectorAll('div.search-result-card__col').forEach((element) => {
        const title = element.querySelector('.item-title__title')?.innerText.trim() || '';
        const description = element.querySelector('.search-result-card__description')?.innerText.trim() || '';
        const idElement = element.querySelector('p.search-result-card__description');
        const id = idElement ? idElement.innerText.trim().replace('ID: ', '') : '';
        const price = element.querySelector('.app-price__amount')?.innerText.trim() || '';

        // Проверяем статус завершенности
        const statusElement = element.querySelector('.search-result-card__label span');
        const completed = statusElement ? statusElement.innerText.trim() : '';

        // Логируем статус для отладки
        console.log(`Status text: "${completed}"`);

        // Если title и description присутствуют, создаем новый объект
        if (title && description) {
          if (currentItem.title) {
            results.push(currentItem); // Добавляем предыдущий объект в результат
          }
          currentItem = { title, description, id, completed, price }; // Создаем новый объект
        } else {
          // Обработка ситуации, когда часть данных (например, цена) может быть в другом месте
          if (price) {
            currentItem.price = price; // Обновляем цену текущего объекта
          }
          if (completed && currentItem.title) {
            currentItem.completed = completed; // Обновляем статус завершенности текущего объекта
          }
        }
      });

      // Добавляем последний объект, если он существует
      if (currentItem.title) {
        results.push(currentItem);
      }

      return results;
    });

    await browser.close();

    // Создаем новый Excel-файл
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');

    // Добавляем заголовки в таблицу
    worksheet.columns = [
      { header: 'Title', key: 'title', width: 30 },
      { header: 'Price', key: 'price', width: 20 },
      { header: 'Description', key: 'description', width: 50 },
      { header: 'ID', key: 'id', width: 30 },
      { header: 'Completed', key: 'completed', width: 25 }
    ];

    // Добавляем данные в Excel-файл
    data.forEach((row, index) => {
      console.log(`Adding row ${index + 1}:`, row); // Для отладки
      worksheet.addRow(row);
    });

    // Сохраняем Excel-файл
    await workbook.xlsx.writeFile('data.xlsx');
    console.log('Data has been saved to data.xlsx');
  } catch (error) {
    console.error('Error:', error.message);
  }
})();
