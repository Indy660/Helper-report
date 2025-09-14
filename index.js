const { Document, Packer, Paragraph, TextRun } = require("docx");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { google } = require("googleapis");

// Авторизация Google API
const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

const spreadsheetId = "1T4rlAKd7z6-6MFGBGVVHxaUaGaw7rgdiYDsv0eCwxCY";

async function getWeeklyReport(targetDate = null) {
    const client = await auth.getClient();
    const sheets = google.sheets({ version: "v4", auth: client });

    // 1. Даты
    const dateRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "A2:A",
    });
    const dates = dateRes.data.values.map(row => row[0]);

    // 2. Имена
    const namesRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "B1:1",
    });
    const names = namesRes.data.values[0];

    // 3. Определяем нужную дату
    let targetRowIndex;
    if (targetDate) {
        targetRowIndex = dates.findIndex(d => d === targetDate);
    } else {
        const today = new Date();
        const todayStr = today.toLocaleDateString("ru-RU", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
        });
        targetRowIndex = dates.findIndex(d => d === todayStr);
    }

    if (targetRowIndex === -1) {
        throw new Error("Дата не найдена в таблице");
    }

    const rowNumber = targetRowIndex + 2;

    // 4. Отчёты за неделю
    const dataRes = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `B${rowNumber}:ZZ${rowNumber}`,
    });
    const reports = dataRes?.data?.values?.[0] || [];

    // 5. Склеиваем имя + отчёт
    const result = names.map((name, i) => ({
        name,
        report: reports[i] || "",
    }));

    const reportDate = dates[targetRowIndex];
    return { result, reportDate };
}

async function createReportDoc(reportData, reportDate) {
    const children = [];

    reportData.forEach(({ name, report }) => {
        // Имя
        children.push(
            new Paragraph({
                children: [new TextRun({ text: name, size: 28 })], // 14pt
                spacing: { after: 200 },
            })
        );

        // Заголовок "Отчет за неделю"
        children.push(
            new Paragraph({
                children: [new TextRun({ text: "Отчет за неделю", size: 28 })], // 14pt
                spacing: { after: 200 },
            })
        );

        if (report.trim() !== "") {
            // Разбиваем задачи
            const tasks = report.split(/\r?\n/).map(t => t.trim()).filter(Boolean);
            tasks.forEach(task => {
                children.push(
                    new Paragraph({
                        children: [new TextRun({ text: task, size: 24 })], // 12pt
                        bullet: { level: 0 },
                    })
                );
            });
        } else {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: "Нет отчета", size: 24, italics: true })],
                })
            );
        }

        // Отступ между людьми (2 энтера)
        children.push(new Paragraph({ text: "" }));
        children.push(new Paragraph({ text: "" }));
    });

    const doc = new Document({
        sections: [{ properties: {}, children }],
    });

    // Название файла
    const formattedDate = reportDate.replace(/\./g, "_");
    const fileName = `Отчет о проделанной работе ${formattedDate}.docx`;

    // Путь: Desktop/Отчеты о проделанной работе
    const desktopPath = path.join(os.homedir(), "Desktop", "Отчеты о проделанной работе");
    if (!fs.existsSync(desktopPath)) {
        fs.mkdirSync(desktopPath, { recursive: true });
    }
    const filePath = path.join(desktopPath, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    console.log(`Файл создан: ${filePath}`);
}

// 🚀 Запуск
(async () => {
    try {
        const { result, reportDate } = await getWeeklyReport("11.09.2025"); // можно null для текущей недели
        await createReportDoc(result, reportDate);
    } catch (err) {
        console.error("Ошибка:", err.message);
    }
})();
