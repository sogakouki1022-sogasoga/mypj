/**
 * Y.Fマネジメント様 概算見積たたき xlsx 作成
 * 実行: node scripts/create_yf_estimate_xlsx.js
 */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const outDir = path.join(__dirname, '..', 'YFマネジメント（株式会社人事部）');
const outPath = path.join(outDir, '概算見積_たたき.xlsx');

const data = [
  ['株式会社Y.Fマネジメント様 概算見積（たたき）', '', '', '', ''],
  ['作成日：', '2026年2月', '', '', ''],
  [],
  ['フェーズ', '項目・範囲', '内容', '概算（税別）', '備考'],
  ['Phase1', '名刺連携・顧客管理・接点履歴', '名刺(Eight)連携、顧客マスタ（法人番号ベース）、部門別情報、接点（活動）履歴。アプリ設計・構築・初期データ整備支援含む。', '100万円前後', 'Eight連携含む'],
  ['Phase2', '案件/タスク管理・日報', '案件/タスク管理（部門別アプリ分割可）、日報連携。タスクの標準化・権限設計含む。', '100万円前後', ''],
  ['Phase3', '売上・契約・マネフォ連携', '売上・契約情報の取り込み、マネーフォワード連携（要確認）。CSV取込 or API連携。', '要ヒアリング', '契約プラン次第'],
  ['保守・運用支援', '月次保守・軽微な改修', '納品後の保守、軽微な改修、問い合わせ対応。', '別途（月額）', 'オプション'],
  [],
  ['合計（Phase1+2想定）', '', '', '300〜400万円前後', ''],
  [],
  ['【注記】', '詳細要件確定後に正式見積を提出します。本見積は目安であり、範囲変更により変動する場合があります。', '', '', ''],
  ['', 'IT導入補助金の対象となる場合、Y.Fマネジメント様との直接契約で申請可能な場合があります（要確認）。', '', '', '']
];

if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

const ws = XLSX.utils.aoa_to_sheet(data);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, '概算見積');

XLSX.writeFile(wb, outPath);
console.log('Created:', outPath);
