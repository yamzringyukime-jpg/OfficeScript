// Office Script: PatternInput の Total を使って ShiftTable に Repeat/Spot を割り当てる
function main(workbook: ExcelScript.Workbook) {
  // --- 設定（必要に応じてテーブル名・列名を変更） ---
  const patternTableName = "PatternInputTable"; // 手入力テーブル
  const shiftTableName = "ShiftTable";         // シフト一覧テーブル

  const colStartName = "開始時間";
  const colEndName = "終了時間";
  const colTotalName = "Total";
  const colRepeatName = "リピートの回答人数(編集可能)";
  const colSpotName = "スポットの回答人数(編集可能)";

  // --- テーブル取得（存在チェック） ---
  const patternTable = workbook.getTable(patternTableName);
  const shiftTable = workbook.getTable(shiftTableName);
  if (!patternTable) {
    console.log(`Error: テーブル '${patternTableName}' が見つかりません。名前を確認してください。`);
    return;
  }
  if (!shiftTable) {
    console.log(`Error: テーブル '${shiftTableName}' が見つかりません。名前を確認してください。`);
    return;
  }

  // --- PatternInput を読み込み、マップ作成（キー: "開始|終了" -> Total） ---
  const patternRange = patternTable.getRangeBetweenHeaderAndTotal();
  const patternValues = patternRange.getValues(); // 2D array
  // ヘッダ列インデックスを取得
  const patternHeader = patternTable.getHeaderRowRange().getValues()[0] as string[];
  const idxPStart = patternHeader.indexOf(colStartName);
  const idxPEnd = patternHeader.indexOf(colEndName);
  const idxPTotal = patternHeader.indexOf(colTotalName);
  if (idxPStart < 0 || idxPEnd < 0 || idxPTotal < 0) {
    console.log("Error: PatternInputTable の列名が見つかりません。開始時間/終了時間/Total の列名を確認してください。");
    return;
  }

  const patternMap = new Map<string, number>();
  for (let r = 0; r < patternValues.length; r++) {
    const start = String(patternValues[r][idxPStart] ?? "").trim();
    const end = String(patternValues[r][idxPEnd] ?? "").trim();
    const totalRaw = patternValues[r][idxPTotal];
    if (!start || !end) continue;
    const total = (totalRaw === null || totalRaw === "") ? null : Number(totalRaw);
    if (total === null || Number.isNaN(total)) continue;
    const key = `${start}|${end}`;
    patternMap.set(key, total);
  }

  if (patternMap.size === 0) {
    console.log("PatternInput に有効な Total が見つかりません。入力を確認してください。");
    return;
  }

  // --- ShiftTable のヘッダとデータを取得 ---
  const shiftHeader = shiftTable.getHeaderRowRange().getValues()[0] as string[];
  const idxSStart = shiftHeader.indexOf(colStartName);
  const idxSEnd = shiftHeader.indexOf(colEndName);
  const idxSRepeat = shiftHeader.indexOf(colRepeatName);
  const idxSSpot = shiftHeader.indexOf(colSpotName);
  if (idxSStart < 0 || idxSEnd < 0) {
    console.log("Error: ShiftTable に開始時間/終了時間列が見つかりません。列名を確認してください。");
    return;
  }
  // リピート/スポット列が無ければ追加（テーブル列追加）
  if (idxSRepeat < 0) {
    shiftTable.addColumn(-1, undefined, colRepeatName);
  }
  if (idxSSpot < 0) {
    shiftTable.addColumn(-1, undefined, colSpotName);
  }
  // 再取得（列インデックスが変わるため）
  const shiftHeader2 = shiftTable.getHeaderRowRange().getValues()[0] as string[];
  const iStart = shiftHeader2.indexOf(colStartName);
  const iEnd = shiftHeader2.indexOf(colEndName);
  const iRepeat = shiftHeader2.indexOf(colRepeatName);
  const iSpot = shiftHeader2.indexOf(colSpotName);

  const shiftRange = shiftTable.getRangeBetweenHeaderAndTotal();
  const shiftValues = shiftRange.getValues(); // 2D array

  // --- 各行を走査して一致するパターンがあれば割当て ---
  for (let r = 0; r < shiftValues.length; r++) {
    const sStart = String(shiftValues[r][iStart] ?? "").trim();
    const sEnd = String(shiftValues[r][iEnd] ?? "").trim();
    const key = `${sStart}|${sEnd}`;
    if (patternMap.has(key)) {
      const total = patternMap.get(key) as number;
      // 割当ロジック：Repeat = floor(total/3), Spot = total - Repeat
      const repeat = Math.floor(total / 3);
      const spot = total - repeat;
      shiftValues[r][iRepeat] = repeat;
      shiftValues[r][iSpot] = spot;
    } else {
      // 一致しない場合は空欄にする（必要ならコメントアウト）
      // shiftValues[r][iRepeat] = "";
      // shiftValues[r][iSpot] = "";
    }
  }

  // --- 変更をテーブルに書き戻す ---
  shiftRange.setValues(shiftValues);
  console.log("完了: ShiftTable に Repeat/Spot を割り当てました。");
}
