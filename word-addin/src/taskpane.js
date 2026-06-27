// Office JavaScript API が読み込まれるのを待つ
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("markdown-button").onclick = convertToMarkdown;
        document.getElementById("neutral-button").onclick = convertToNeutral;
        document.getElementById("outline-button").onclick = convertToOutline;
    }
});

function showStatus(message, isSuccess = true) {
    const status = document.getElementById("status");
    status.textContent = message;
    status.className = isSuccess ? "success" : "error";
    setTimeout(() => {
        status.className = "";
    }, 3000);
}

async function convertToMarkdown() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.paragraphs.items.forEach((paragraph) => {
                paragraph.leftIndent = 0;
                paragraph.firstLineIndent = 0;
            });
            await context.sync();
            showStatus("✅ マークダウン形式に変換しました", true);
        });
    } catch (error) {
        showStatus("❌ エラー: " + error.message, false);
    }
}

async function convertToNeutral() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.paragraphs.items.forEach((paragraph) => {
                paragraph.leftIndent = 0;
                paragraph.firstLineIndent = 0;
                paragraph.lineSpacing = 15;
            });
            await context.sync();
            showStatus("✅ ニュートラル形式に変換しました", true);
        });
    } catch (error) {
        showStatus("❌ エラー: " + error.message, false);
    }
}

async function convertToOutline() {
    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            body.paragraphs.items.forEach((paragraph) => {
                const level = paragraph.outlineLevel || 0;
                paragraph.leftIndent = level * 720;
            });
            await context.sync();
            showStatus("✅ アウトライン形式に変換しました", true);
        });
    } catch (error) {
        showStatus("❌ エラー: " + error.message, false);
    }
}
