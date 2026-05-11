//Patternの開始時間と終了時間をPatternInputに入力

function main(workbook: ExcelScript.Workbook) {
  // --- 設定（テーブル名） ---
  const patternTableName = "PatternTable";
  const inputTableName = "PatternInputTable";

  const patternTable = workbook.getTable(patternTableName);
  const inputTable = workbook.getTable(inputTableName);

  if (!patternTable) {
    console.log(`Error: テーブル '${patternTableName}' が見つかりません。`);
    return;
  }
  if (!inputTable) {
    console.log(`Error: テーブル '${inputTableName}' が見つかりません。`);
    return;
  }

  // --- 1. PatternInputTable の既存の Total を記憶 ---
  const inputMap = new Map<string, string | number | boolean>();
  const inputRowCount = inputTable.getRowCount();
  const inputHeaders = inputTable.getHeaderRowRange().getValues()[0] as string[];

  // 列のインデックスを動的に取得
  const iStart = inputHeaders.indexOf("開始時間");
  const iEnd = inputHeaders.indexOf("終了時間");
  const iTotal = inputHeaders.indexOf("Total");

  if (iStart < 0 || iEnd < 0 || iTotal < 0) {
    console.log("Error: PatternInputTable に開始時間/終了時間/Totalの列が見つかりません。");
    return;
  }

  // 既存データがあれば、"開始時間|終了時間" をキーにして Total を保存
  if (inputRowCount > 0) {
    const inputValues = inputTable.getRangeBetweenHeaderAndTotal().getValues();
    for (let r = 0; r < inputValues.length; r++) {
      const start = String(inputValues[r][iStart] ?? "").trim();
      const end = String(inputValues[r][iEnd] ?? "").trim();
      const total = inputValues[r][iTotal];
      // Totalが空欄でなければ保存
      if (start && end && total !== "") {
        inputMap.set(`${start}|${end}`, total);
      }
    }
  }

  // --- 2. PatternTable から最新の 開始時間・終了時間 を取得 ---
  const patternHeaders = patternTable.getHeaderRowRange().getValues()[0] as string[];
  const pStart = patternHeaders.indexOf("開始時間");
  const pEnd = patternHeaders.indexOf("終了時間");

  if (pStart < 0 || pEnd < 0) {
      console.log("Error: PatternTable に開始時間/終了時間の列が見つかりません。");
      return;
  }

  const pRowCount = patternTable.getRowCount();
  if (pRowCount === 0) {
    console.log("PatternTable にデータがありません。Power Queryの更新を確認してください。");
    return;
  }

  const pValues = patternTable.getRangeBetweenHeaderAndTotal().getValues();
  let newData: (string | number | boolean)[][] = [];

  // --- 3. 書き込むデータ配列を作成（保存した Total があれば復元する） ---
  for (let r = 0; r < pValues.length; r++) {
    const start = String(pValues[r][pStart] ?? "").trim();
    const end = String(pValues[r][pEnd] ?? "").trim();

    if (!start && !end) continue; // 空行スキップ

    const key = `${start}|${end}`;
    // エラー対策：引き継ぐTotalの型を明記
    const savedTotal: string | number | boolean = inputMap.has(key) ? (inputMap.get(key) as string | number | boolean) : "";

    // エラー対策：配列(row)の型を明記
    let row: (string | number | boolean)[] = new Array(inputHeaders.length).fill("");
    row[iStart] = start;
    row[iEnd] = end;
    row[iTotal] = savedTotal;

    newData.push(row);
  }

  // --- 4. PatternInputTable を更新 ---
  // 既存のデータ行をすべて削除（行数が0より大きい場合のみ）
  if (inputRowCount > 0) {
    inputTable.deleteRowsAt(0, inputRowCount);
  }

  // 新しいデータを一括追加
  if (newData.length > 0) {
    inputTable.addRows(-1, newData);
  }

  console.log("完了: PatternTable の開始時間・終了時間を PatternInputTable に反映しました。");
}