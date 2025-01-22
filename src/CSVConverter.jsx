import React, { useState } from 'react';
import { read, utils, write } from 'xlsx';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';
//import './CSVConverter.css';

const CSVConverter = () => {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(false);
  const [convertedData, setConvertedData] = useState(null);

  // CSVデータを変換する関数
  const convertCSVData = (csvData) => {
    const converted = csvData.map(row => ({
      '受注番号': row['注文番号'],
      '受注日': row['注文日'],
      '発送予定日': '',
      '出荷予定日': '',
      '配送業者': '',
      '送付先郵便番号': row['お届け先郵便番号'],
      '送付先住所都道府県': row['お届け先住所１'],
      '送付先住所郡市区': row['お届け先住所２'],
      '送付先住所それ以降': row['お届け先住所３'],
      '送付先姓': row['お届け先名称１'],
      '送付先名': row['お届け先名称２'],
      '送付先電話番号': row['お届け先電話番号'],
      '商品コード': '',
      '商品名': '',
      '商品オプション': '',
      '数量': row['出荷個数'],
      '単価': '',
      'オプション価格': '',
      '消費税率': '',
      '代引料': '',
      '送料': '',
      '手数料': '',
      'ポイント利用額': '',
      'その他費用': '',
      '合計金額': '',
      'ギフトフラグ': '',
      '時間帯指定': row['配達指定時間帯'],
      '日付指定': row['配達日'],
      '作業者欄': '',
      '備考': ''
    }));
    return converted;
  };

  // バイナリ文字列をArrayBufferに変換
  const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  };

  // Excelファイルとしてダウンロードする関数
  const downloadExcel = (data) => {
    const ws = utils.json_to_sheet(data);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "変換データ");

    const fileName = "converted_data.xlsx";
    const wbout = write(wb, { bookType: 'xlsx', type: 'binary' });
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), fileName);
  };

  // ファイルアップロード処理
  const handleFileUpload = async (event) => {
    try {
      setLoading(true);
      setError(null);
      setSuccess(false);

      const file = event.target.files[0];
      if (!file) {
        throw new Error('ファイルが選択されていません。');
      }

      if (!file.name.endsWith('.csv')) {
        throw new Error('CSVファイルを選択してください。');
      }

      // handleFileUpload関数内のPapa.parseの呼び出しを修正
      Papa.parse(file, {
        complete: (results) => {
          try {
            if (results.errors.length > 0) {
              throw new Error('CSVファイルの解析中にエラーが発生しました。');
            }

            const converted = convertCSVData(results.data);
            setConvertedData(converted);
            downloadExcel(converted);
            setSuccess(true);
          } catch (err) {
            setError(err.message);
          }
        },
        header: true,
        encoding: 'Shift-JIS', // 文字エンコーディングを指定
        error: (error) => {
          setError('ファイルの読み込み中にエラーが発生しました。');
        }
      });
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="converter-container">
      <h1 className="converter-title">CSV変換ツール</h1>

      <div className="upload-section">
        <label className="file-upload-label">
          <input
            type="file"
            accept=".csv"
            onChange={handleFileUpload}
            disabled={loading}
            className="file-input"
          />
          <span className="upload-button">
            {loading ? 'アップロード中...' : 'CSVファイルを選択'}
          </span>
        </label>
      </div>

      {error && (
        <div className="error-message">
          <p>エラー: {error}</p>
        </div>
      )}

      {success && (
        <div className="success-message">
          <p>変換が完了しました！</p>
        </div>
      )}

      <div className="instructions">
        <h2>使用方法</h2>
        <ol>
          <li>変換したいCSVファイルを選択してください。</li>
          <li>自動的に変換が開始されます。</li>
          <li>変換されたExcelファイルが自動的にダウンロードされます。</li>
        </ol>
      </div>
    </div>
  );
};

export default CSVConverter;