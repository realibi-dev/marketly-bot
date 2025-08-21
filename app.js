const TelegramBot = require('node-telegram-bot-api');
const moment = require('moment');
const fs = require('fs');
const writeXlsxFile = require('write-excel-file/node')
const pg = require('pg');
const path = require('path');
const { Client } = pg
const token = '8347936739:AAF_kUuXhLUrKJGUeWkGNt2Be3asIDAZdt8';

const client = new Client({
  user: 'postgres',
  password: 'admin',
  host: '194.32.140.16',
  port: 5432,
  database: 'postgres',
});

function handleNakCommand(bot, chatId, providers) {
  const inlineKeyboard = {
    reply_markup: {
      inline_keyboard: providers.map(provider => ([{ text: provider.name, callback_data: `consignment ${provider.id}` }]))
    }
  };

  bot.sendMessage(chatId, 'Выберите поставщика', inlineKeyboard);
}

async function handleConsignmentCommand(bot, chatId, provider, client) {
  try {
    const { rows: allOrderProducts } = await client.query(`
        select product.name             as "productName",
               "orderItem".quantity,
               branch."name"            as "branchName",
               branch.id                as "branchId",
               branch.address           as "branchAddress",
               "providerProfile"."name" as "providerName",
               "providerProfile".id     as "providerId",
               "order"."orderNumber"               as "order_number",
               "order"."createdAt",
               "orderItem".price            as "price"
        from "orderItem"
                 left join product on product.id = "orderItem"."productId"
                 left join "order" on "orderItem"."orderId" = "order".id
                 left join branch on branch.id = "order"."branchId"
                 left join "providerProfile" on "providerProfile".id = product."providerId"
        where product."providerId" = $1
          AND "order"."isCompleted" is true
          AND "order"."createdAt" >= CURRENT_DATE::timestamp 
          AND "order"."createdAt" < (CURRENT_DATE + INTERVAL '1 day')::timestamp
          AND "order"."deletedAt" is null
    `, [provider.id]);

    console.log(`Найдено ${allOrderProducts.length} продуктов для поставщика ${provider.name}`);

    if (allOrderProducts.length === 0) {
      bot.sendMessage(chatId, `На сегодня нет заказов для поставщика "${provider.name}"`);
      return;
    }

    const uniqueBranchIds = [...new Set(allOrderProducts.map(p => p.branchId))];

    // Создаем массивы для хранения данных всех листов
    const EXCELL_TOTAL_SHEETS = [];
    const EXCELL_TOTAL_SHEETS_NAMES = [];
    const EXCELL_TOTAL_COLUMNS = [];

    for (const branchId of uniqueBranchIds) {
      const branchProducts = allOrderProducts.filter(product => product.branchId === branchId);

      const EXCELL_ROWS_ARRAY = [
        [
          {
            value: 'Номер заказа',
            fontWeight: 'bold',
          },
          {
            value: branchProducts[0].order_number || 'N/A',
          }
        ],
        [
          {
            value: 'Ресторан',
            fontWeight: 'bold',
          },
          {
            value: `${branchProducts[0].branchName} - ${branchProducts[0].branchAddress}`,
          }
        ],
        [
          {
            value: 'Дата создания',
            fontWeight: 'bold',
          },
          {
            value: moment(branchProducts[0].createdAt).format("YYYY-MM-DD"),
          }
        ],
        [null, null, null, null, null],
        [
          {
            value: '#',
            fontWeight: 'bold',
          },
          {
            value: 'Наименование',
            fontWeight: 'bold',
          },
          {
            value: 'Количество',
            fontWeight: 'bold',
          },
          {
            value: 'Цена',
            fontWeight: 'bold',
          },
          {
            value: 'Сумма',
            fontWeight: 'bold',
          },
        ]
      ];

      let totalBranchProductsCount = 0;
      let totalBranchProductsPrice = 0;
      let orderNumber = 1;

      for (const product of branchProducts) {
        const price = parseFloat(product.price) || 0;
        const quantity = parseInt(product.quantity) || 0;
        const sum = price * quantity;

        EXCELL_ROWS_ARRAY.push([
          {
            value: orderNumber,
          },
          {
            value: product.productName || '',
          },
          {
            value: quantity,
          },
          {
            value: price,
          },
          {
            value: sum,
          },
        ]);

        totalBranchProductsCount += quantity;
        totalBranchProductsPrice += sum;
        orderNumber++;
      }

      EXCELL_ROWS_ARRAY.push([
        {
          value: '',
        },
        {
          value: 'Итого',
          fontWeight: 'bold',
        },
        {
          value: totalBranchProductsCount,
          fontWeight: 'bold',
        },
        {
          value: '',
        },
        {
          value: Math.round(totalBranchProductsPrice * 100) / 100, // Округляем до копеек
          fontWeight: 'bold',
        },
      ]);

      // Создаем имя листа в формате [branch.name] [branch.address]
      const sheetName = branchProducts[0].branchName;

      // Добавляем данные в массивы
      EXCELL_TOTAL_SHEETS.push(EXCELL_ROWS_ARRAY);
      EXCELL_TOTAL_SHEETS_NAMES.push(sheetName);
      EXCELL_TOTAL_COLUMNS.push([
        { width: 15 }, // #
        { width: 50 }, // Наименование
        { width: 12 }, // Количество
        { width: 15 }, // Цена
        { width: 15 }, // Сумма
      ]);
    }

    // Создаем один файл со всеми листами
    const fileName = `Накладные_${provider.name.replace(/[:"]/g, '').replace(/\s+/g, '_')}_${moment().format("YYYY_MM_DD")}_${Math.round(Math.random() * 100000)}.xlsx`;
    const excellFilePath = path.join(__dirname, fileName);

    try {
      // Создаем Excel файл с несколькими листами
      await writeXlsxFile(
        EXCELL_TOTAL_SHEETS,
        {
          sheets: EXCELL_TOTAL_SHEETS_NAMES,
          columns: EXCELL_TOTAL_COLUMNS,
          filePath: excellFilePath,
        }
      );

      console.log(`Excel файл с ${EXCELL_TOTAL_SHEETS.length} листами создан: ${excellFilePath}`);

      const stats = fs.statSync(excellFilePath);
      console.log(`Размер файла: ${stats.size} байт`);

      const fileBuffer = fs.readFileSync(excellFilePath);

      await bot.sendDocument(chatId, fileBuffer, {}, {
        filename: fileName,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      // Удаляем временный файл
      fs.unlinkSync(excellFilePath);

    } catch (error) {
      console.error('Ошибка при создании Excel файла:', error);
      bot.sendMessage(chatId, `Ошибка при создании файла с накладными`);
    }

  } catch (error) {
    console.error('Ошибка в handleConsignmentCommand:', error);
    bot.sendMessage(chatId, 'Произошла ошибка при обработке запроса');
  }
}

async function handleTotalDailyCommand(bot, chatId, client, msg) {
  const { rows: orderedProducts } = await client.query(`
    select
    distinct p.name as product_name,
    COALESCE((
      select sum(oi.quantity)
      from "orderItem" oi
        join "order" oo on oo.id = oi."orderId"
        where oi."productId" = p.id and
        oo."createdAt" >= CURRENT_DATE::timestamp and
        oo."createdAt" < (CURRENT_DATE + INTERVAL '1 day')::timestamp and
        oo."isCompleted" is true AND
        oo."deletedAt" is null
    ), 0) as "count",
    pr."name" as provider_name,
    pr.id as provider_id,
    p.id as product_id
    from product p
    left join "providerProfile" pr on pr.id = p."providerId"
    left join "orderItem" oi on oi."productId" = p.id
    left join "order" oo on oo.id = oi."orderId"
    where
    oo."deletedAt" is null AND
    oo."isCompleted" is true AND
    oo."createdAt" >= CURRENT_DATE::timestamp and
    oo."createdAt" < (CURRENT_DATE + INTERVAL '1 day')::timestamp
  `);

  let providers = [...new Set(orderedProducts.map(op => op.provider_id))]

  if (!providers.length) {
    bot.sendMessage(chatId, "На данный момент нет подтвержденных заказов на завтра");
    return;
  }

  providers = providers.map(providerId => ({
    id: providerId,
    name: orderedProducts.find(op => op.provider_id === providerId).provider_name,
  }));

  for (const provider of providers) {
    const providerProducts = orderedProducts.filter(product => product.provider_id == provider.id);
    let message = `Заявка на ${moment().add(1, 'days').format('DD/MM/YYYY')}.\n${provider.name}\n`;
    for (const product of providerProducts) {
      message = message.concat(`\n${product.product_name} - ${product.count} шт.`);
    }

    bot.sendMessage(chatId, message);
  }

  console.log(`[${moment().format('DD-MM-YYYY HH:m:ss')}] Отправлен итог на сегодня`);
}

client.connect().then(async () => {
  const bot = new TelegramBot(token, { polling: true });

  const { rows: activeProviders } = await client.query(`
      select *
      from "providerProfile"
      where "isActive" = true
        and "deletedAt" is NULL
  `);

  bot.onText(/\/start/, (msg) => {
    const chatId = msg.chat.id;

    const inlineKeyboard = {
      reply_markup: {
        inline_keyboard: [
          [
            { text: 'Итог на сегодня', callback_data: 'total daily' },
            { text: 'Накладные для склада', callback_data: 'nak' },
          ],
        ]
      }
    };

    bot.sendMessage(chatId, 'Привет! Что ты хочешь получить?', inlineKeyboard);
  });

  // ✅ Исправлено - используем обычные кнопки без callback_data
  bot.onText(/\/menu/, (msg) => {
    const chatId = msg.chat.id;

    const replyKeyboard = {
      reply_markup: {
        keyboard: [
          [
            { text: 'Итог на сегодня' },
            { text: 'Накладные для склада' },
          ],
        ],
        resize_keyboard: true,
        one_time_keyboard: false
      }
    };

    bot.sendMessage(chatId, 'Что ты хочешь получить?', replyKeyboard);
  });

  bot.onText(/Накладные для склада/, (msg) => {
    const chatId = msg.chat.id;
    handleNakCommand(bot, chatId, activeProviders);
  });

  bot.onText(/Итог на сегодня/, (msg) => {
    const chatId = msg.chat.id;
    handleTotalDailyCommand(bot, chatId, client);
  });

  // ✅ Исправлено - используем data вместо message
  bot.on('callback_query', (callbackQuery) => {
    const message = callbackQuery.message;
    const data = callbackQuery.data; // ← Используем data
    const chatId = message.chat.id;

    if (data.startsWith('consignment')) { // ← Исправлено
      const providerId = +data.split(" ")[1]; // ← Исправлено
      const provider = activeProviders.find(provider => provider.id === providerId);
      handleConsignmentCommand(bot, chatId, provider, client);
    } else if (data.startsWith('nak')) { // ← Исправлено
      handleNakCommand(bot, chatId, activeProviders);
    } else if (data.startsWith('total daily')) { // ← Исправлено
      handleTotalDailyCommand(bot, chatId, client, message);
    } else {
      bot.answerCallbackQuery(callbackQuery.id, `You pressed: ${data}`);
    }

    // Always answer the callback query to remove loading state
    bot.answerCallbackQuery(callbackQuery.id);
  });
});