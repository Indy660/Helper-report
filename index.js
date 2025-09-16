const { Document, Packer, Paragraph, TextRun } = require("docx");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { google } = require("googleapis");

require("dotenv").config();

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
    const fileName = `Отчет о проделанной работе ${reportDate}.docx`;

    // 📁 Файловая структура: Desktop/Отчеты о проделанной работе/[дата]
    const folderName = reportDate; // например "11.09.2025"
    const desktopPath = path.join(os.homedir(), "Desktop", "Отчеты о проделанной работе", folderName);
    if (!fs.existsSync(desktopPath)) {
        fs.mkdirSync(desktopPath, { recursive: true });
    }
    const filePath = path.join(desktopPath, fileName);

    // Путь: Desktop/Отчеты о проделанной работе
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    console.log(`Файл создан: ${filePath}`);
}

//////////////////////////////////

const OpenAI = require("openai");

const client = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
});

async function createAIReport(reportData, reportDate) {
    // Подготавливаем текст для AI
    let textForAI = `Ниже представлены отчёты сотрудников за неделю ${reportDate}:\n\n`;
    reportData.forEach(({ name, report }) => {
        textForAI += `${report || "нет отчёта"}\n\n`;
    });

    // Запрос в AI
    const completion = await client.chat.completions.create({
        model: "gpt-4o-mini",
        messages: [
            {
                role: "system",
                content: "Ты помощник-аналитик, составляешь структурированные еженедельные отчёты.",
            },
            {
                role: "user",
                content: `Составь единый "Общий отчет за неделю" в следующем формате:
1. Заголовок "Общий отчет за неделю" (жирный, 14pt).
2. Краткое вводное предложение о проделанной работе.
3. Раздел "Основные выполненные задачи" (жирный, 14pt), далее список:
   • пункт первого уровня
   o подпункты второго уровня (если есть)
   Не указывай имена сотрудников, перечисляй только выполненные задачи.
4. Раздел "Внедрение новых и сопровождение текущих сотрудников" (жирный, 14pt).
5. Раздел "Глобальное обновление проектов" (жирный, 14pt).
6. Раздел "Выводы" (жирный, 14pt) — 2–8 тезиса в формате ✅.
7. В конце добавь строку:
   "План на ближайшее время: улучшение кодовой базы, исправление ошибок и оптимизация работы приложений."

Не используй Markdown и заголовки с #. Просто текст с перечислениями.

Вот отчёты сотрудников:\n\n${textForAI}`,
            },
        ],
    });

    const analysis = completion.choices[0].message.content;

    // Формируем Word документ
    const children = [];
    analysis.split(/\r?\n/).forEach(line => {
        if (!line.trim()) {
            children.push(new Paragraph({ text: "" }));
        } else if (
            line.startsWith("Общий отчет") ||
            line.startsWith("Основные выполненные задачи") ||
            line.startsWith("Внедрение новых") ||
            line.startsWith("Глобальное обновление проектов") ||
            line.startsWith("Выводы")
        ) {
            // Заголовки разделов — жирные, 14pt
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: line.trim(), size: 28, bold: true })],
                    spacing: { after: 200 },
                })
            );
        } else if (line.startsWith("•")) {
            // Буллет первого уровня
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: line.replace(/^•\s*/, ""), size: 24 })],
                    bullet: { level: 0 },
                })
            );
        } else if (line.startsWith("o")) {
            // Подпункт второго уровня
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: line.replace(/^o\s*/, ""), size: 24 })],
                    bullet: { level: 1 },
                })
            );
        } else if (line.startsWith("✅")) {
            // Выводы — чекбокс
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: line, size: 24 })],
                })
            );
        } else {
            // Обычный текст
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: line, size: 24 })],
                })
            );
        }
    });

    const doc = new Document({
        sections: [{ properties: {}, children }],
    });

    // 📁 Файловая структура: Desktop/Отчеты о проделанной работе/[дата]
    const folderName = reportDate; // например "11.09.2025"
    const desktopPath = path.join(os.homedir(), "Desktop", "Отчеты о проделанной работе", folderName);
    if (!fs.existsSync(desktopPath)) {
        fs.mkdirSync(desktopPath, { recursive: true });
    }

    const fileName = `Общий отчет за неделю для руководства ${reportDate}.docx`;
    const filePath = path.join(desktopPath, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    console.log(`Общий отчет создан: ${filePath}`);
}

module.exports = { createAIReport };

// 🚀 Запуск
(async () => {
    try {
        const { result, reportDate } = await getWeeklyReport("11.09.2025"); // можно null для текущей недели
        await createReportDoc(result, reportDate);
        await createAIReport(result, reportDate);
    } catch (err) {
        console.error("Ошибка:", err.message);
    }
})();
