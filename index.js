const { google } = require("googleapis");
const fs = require("fs");

// Авторизация Google API
const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

const spreadsheetId = "1T4rlAKd7z6-6MFGBGVVHxaUaGaw7rgdiYDsv0eCwxCY";

async function getWeeklyReport(targetDate = null) {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // 1. Читаем все даты (столбец А)
    const dateRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "A2:A", // все даты
    });
    const dates = dateRes.data.values.map(row => row[0]);

    // 2. Читаем имена (строка 1, начиная с B)
    const namesRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "B1:1",
    });
    const names = namesRes.data.values[0];

    // 3. Определяем нужную дату
    let targetRowIndex;
    if (targetDate) {
        // ищем точное совпадение
        targetRowIndex = dates.findIndex(d => d === targetDate);
    } else {
        // ищем ближайший четверг к сегодня
        const today = new Date();
        const todayStr = today.toLocaleDateString("ru-RU", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
        });
        targetRowIndex = dates.findIndex(d => d === todayStr);
    }

    if (targetRowIndex === -1) {
        console.log("Дата не найдена в таблице");
        return;
    }

    // 4. Читаем строку с результатами для этой даты
    const rowNumber = targetRowIndex + 2; // +2 потому что A2 = первая дата
    const dataRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `B${rowNumber}:ZZ${rowNumber}`, // на всякий случай до конца строки
    });
    const reports = dataRes?.data?.values?.[0] || [];

    // 5. Склеиваем "Имя → что сделал"
    const result = names.map((name, i) => ({
        name,
        report: reports[i] || "—",
    }));

    return result;
}

// Пример использования
(async () => {
    // По умолчанию — за текущую неделю
    const thisWeek = await getWeeklyReport();
    console.log("Отчёт за эту неделю:");
    console.table(thisWeek);

    // Или за конкретную дату
    const specific = await getWeeklyReport("10.09.2025");
    console.log("Отчёт за 11.09.2025:");
    console.table(specific);
})();