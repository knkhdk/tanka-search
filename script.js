let workbook = null;
let worksheet = null;

// 数値をカンマ区切りの文字列に変換する関数
function formatNumber(num) {
    if (num === undefined || num === null || num === '') return '';
    return Number(num).toLocaleString('ja-JP');
}

// ページ読み込み時に自動的にExcelファイルを読み込む
window.addEventListener('load', function() {
    fetch('tanka.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            workbook = XLSX.read(new Uint8Array(data), { type: 'array' });
            worksheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // デバッグ用：読み込んだデータの内容を確認
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            console.log('読み込んだデータ:', jsonData);
            console.log('列名:', Object.keys(jsonData[0] || {}));
            
            // ファイルが読み込まれたことを通知
            alert('単価データを読み込みました。検索を開始できます。\n列名: ' + Object.keys(jsonData[0] || {}).join(', '));
        })
        .catch(error => {
            console.error('Excelファイルの読み込みに失敗しました:', error);
            alert('単価データの読み込みに失敗しました。tanka.xlsxファイルが正しい場所にあることを確認してください。');
        });
});

// エンターキーでの検索実行を追加
document.getElementById('searchInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        searchWord();
    }
});

function searchWord() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const resultsBody = document.getElementById('resultsBody');
    resultsBody.innerHTML = '';

    if (!worksheet) {
        alert('単価データの読み込みに失敗しています。ページを更新してください。');
        return;
    }

    if (!searchTerm) {
        alert('検索語を入力してください。');
        return;
    }

    // Excelデータを配列に変換
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    let found = false;

    // デバッグ用：検索語を表示
    console.log('検索語:', searchTerm);

    // 検索実行
    jsonData.forEach(row => {
        // デバッグ用：各行のデータを表示
        console.log('検索中の行:', row);
        
        // オブジェクトの各プロパティを検索
        Object.entries(row).forEach(([key, value]) => {
            const stringValue = String(value).toLowerCase();
            if (stringValue.includes(searchTerm)) {
                found = true;
                const tr = document.createElement('tr');
                
                // 商品名、単価、備考の列を表示（列名が異なる場合は適宜調整）
                // 単価はカンマ区切りで表示
                tr.innerHTML = `
                    <td>${row['商品名'] || row['名称'] || row['品名'] || ''}</td>
                    <td>${formatNumber(row['単価'] || row['価格'] || row['金額'] || '')}</td>
                    <td>${row['備考'] || row['メモ'] || row['注記'] || ''}</td>
                `;
                
                resultsBody.appendChild(tr);
            }
        });
    });

    if (!found) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td colspan="3" style="text-align: center;">該当するデータが見つかりませんでした。</td>';
        resultsBody.appendChild(tr);
    }
} 