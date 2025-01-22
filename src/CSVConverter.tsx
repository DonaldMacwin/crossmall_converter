import React, { useState, ChangeEvent } from 'react';
import { read, utils, write } from 'xlsx';
import Papa from 'papaparse';
import { saveAs } from 'file-saver';

interface CSVRow {
  '注文番号': string;
  '注文日': string;
  'お届け先郵便番号': string;
  'お届け先住所１': string;
  'お届け先住所２': string;
  'お届け先住所３': string;
  'お届け先名称１': string;
  'お届け先名称２': string;
  'お届け先電話番号': string;
  '出荷個数': string;
  '配達指定時間帯': string;
  '配達日': string;
  [key: string]: string; // インデックスシグネチャ
}

interface ConvertedRow {
  '注文番号': string;
  'ステータス': string;
  'サブステータスID': string;
  'サブステータス': string;
  '注文日時': string;
  '注文日': string;
  '注文時間': string;
  'キャンセル期限日': string;
  '注文確認日時': string;
  '注文確定日時': string;
  '発送指示日時': string;
  '発送完了報告日時': string;
  '支払方法名': string;
  'クレジットカード支払い方法': string;
  'クレジットカード支払い回数': string;
  '配送方法': string;
  '配送区分': string;
  '注文種別': string;
  '複数送付先フラグ': string;
  '送付先一致フラグ': string;
  '離島フラグ': string;
  '楽天確認中フラグ': string;
  '警告表示タイプ': string;
  '楽天会員フラグ': string;
  '購入履歴修正有無フラグ': string;
  '商品合計金額': string;
  '消費税合計': string;
  '送料合計': string;
  '代引料合計': string;
  '請求金額': string;
  '合計金額': string;
  'ポイント利用額': string;
  'クーポン利用総額': string;
  '店舗発行クーポン利用額': string;
  '楽天発行クーポン利用額': string;
  '注文者郵便番号1': string;
  '注文者郵便番号2': string;
  '注文者住所都道府県': string;
  '注文者住所郡市区': string;
  '注文者住所それ以降の住所': string;
  '注文者姓': string;
  '注文者名': string;
  '注文者姓カナ': string;
  '注文者名カナ': string;
  '注文者電話番号1': string;
  '注文者電話番号2': string;
  '注文者電話番号3': string;
  '注文者メールアドレス': string;
  '注文者性別': string;
  '申込番号': string;
  '申込お届け回数': string;
  '送付先ID': string;
  '送付先送料': string;
  '送付先代引料': string;
  '送付先消費税合計': string;
  '送付先商品合計金額': string;
  '送付先合計金額': string;
  'のし': string;
  '送付先郵便番号1': string;
  '送付先郵便番号2': string;
  '送付先住所都道府県': string;
  '送付先住所郡市区': string;
  '送付先住所それ以降の住所': string;
  '送付先姓': string;
  '送付先名': string;
  '送付先姓カナ': string;
  '送付先名カナ': string;
  '送付先電話番号1': string;
  '送付先電話番号2': string;
  '送付先電話番号3': string;
  '商品明細ID': string;
  '商品ID': string;
  'シンセー商品名': string;
  'マスター品番': string;
  '商品管理番号': string;
  '楽天単価': string;
  '個数': string;
  '送料込別': string;
  '税込別': string;
  '代引手数料込別': string;
  '項目・選択肢': string;
  'ポイント倍率': string;
  '納期情報': string;
  '在庫タイプ': string;
  'ラッピングタイトル1': string;
  'ラッピング名1': string;
  'ラッピング料金1': string;
  'ラッピング税込別1': string;
  'ラッピング種類1': string;
  'ラッピングタイトル2': string;
  'ラッピング名2': string;
  'ラッピング料金2': string;
  'ラッピング税込別2': string;
  'ラッピング種類2': string;
  'お届け時間帯': string;
  'お届け日指定': string;
  '担当者': string;
  'ひとことメモ': string;
  'メール差込文': string;
  'ギフト配送希望': string;
  'コメント': string;
  '利用端末': string;
  'メールキャリアコード': string;
  'あす楽希望フラグ': string;
  '医薬品受注フラグ': string;
  '楽天スーパーDEAL商品受注フラグ': string;
  'メンバーシッププログラム受注タイプ': string;
  '決済手数料': string;
  '注文者負担金合計': string;
  '店舗負担金合計': string;
}

const CSVConverter: React.FC = () => {
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<boolean>(false);
  const [convertedData, setConvertedData] = useState<ConvertedRow[] | null>(null);

  // CSVデータを変換する関数
  const convertCSVData = (csvData: CSVRow[]): ConvertedRow[] => {
    return csvData.map(row => {
      // 郵便番号を分割（xxx-xxxx の形式を想定）
      const postalCode = row['お届け先郵便番号'].split('-');
      const postalCode1 = postalCode[0] || '';
      const postalCode2 = postalCode[1] || '';
  
      // 電話番号を分割（xxx-xxxx-xxxx の形式を想定）
      const phone = row['お届け先電話番号'].split('-');
      const phone1 = phone[0] || '';
      const phone2 = phone[1] || '';
      const phone3 = phone[2] || '';
  
      return {
        '注文番号': row['注文番号'],
        'ステータス': '',
        'サブステータスID': '',
        'サブステータス': '',
        '注文日時': row['注文日'],
        '注文日': row['注文日'],
        '注文時間': '',
        'キャンセル期限日': '',
        '注文確認日時': '',
        '注文確定日時': '',
        '発送指示日時': '',
        '発送完了報告日時': '',
        '支払方法名': '',
        'クレジットカード支払い方法': '',
        'クレジットカード支払い回数': '',
        '配送方法': '',
        '配送区分': '',
        '注文種別': '',
        '複数送付先フラグ': '',
        '送付先一致フラグ': '',
        '離島フラグ': '',
        '楽天確認中フラグ': '',
        '警告表示タイプ': '',
        '楽天会員フラグ': '',
        '購入履歴修正有無フラグ': '',
        '商品合計金額': '',
        '消費税合計': '',
        '送料合計': '',
        '代引料合計': '',
        '請求金額': '',
        '合計金額': '',
        'ポイント利用額': '',
        'クーポン利用総額': '',
        '店舗発行クーポン利用額': '',
        '楽天発行クーポン利用額': '',
        '注文者郵便番号1': postalCode1,
        '注文者郵便番号2': postalCode2,
        '注文者住所都道府県': row['お届け先住所１'],
        '注文者住所郡市区': row['お届け先住所２'],
        '注文者住所それ以降の住所': row['お届け先住所３'],
        '注文者姓': row['お届け先名称１'],
        '注文者名': row['お届け先名称２'],
        '注文者姓カナ': '',
        '注文者名カナ': '',
        '注文者電話番号1': phone1,
        '注文者電話番号2': phone2,
        '注文者電話番号3': phone3,
        '注文者メールアドレス': '',
        '注文者性別': '',
        '申込番号': '',
        '申込お届け回数': '',
        '送付先ID': '',
        '送付先送料': '',
        '送付先代引料': '',
        '送付先消費税合計': '',
        '送付先商品合計金額': '',
        '送付先合計金額': '',
        'のし': '',
        '送付先郵便番号1': postalCode1,
        '送付先郵便番号2': postalCode2,
        '送付先住所都道府県': row['お届け先住所１'],
        '送付先住所郡市区': row['お届け先住所２'],
        '送付先住所それ以降の住所': row['お届け先住所３'],
        '送付先姓': row['お届け先名称１'],
        '送付先名': row['お届け先名称２'],
        '送付先姓カナ': '',
        '送付先名カナ': '',
        '送付先電話番号1': phone1,
        '送付先電話番号2': phone2,
        '送付先電話番号3': phone3,
        '商品明細ID': '',
        '商品ID': '',
        'シンセー商品名': '',
        'マスター品番': '',
        '商品管理番号': '',
        '楽天単価': '',
        '個数': row['出荷個数'],
        '送料込別': '',
        '税込別': '',
        '代引手数料込別': '',
        '項目・選択肢': '',
        'ポイント倍率': '',
        '納期情報': '',
        '在庫タイプ': '',
        'ラッピングタイトル1': '',
        'ラッピング名1': '',
        'ラッピング料金1': '',
        'ラッピング税込別1': '',
        'ラッピング種類1': '',
        'ラッピングタイトル2': '',
        'ラッピング名2': '',
        'ラッピング料金2': '',
        'ラッピング税込別2': '',
        'ラッピング種類2': '',
        'お届け時間帯': row['配達指定時間帯'],
        'お届け日指定': row['配達日'],
        '担当者': '',
        'ひとことメモ': '',
        'メール差込文': '',
        'ギフト配送希望': '',
        'コメント': '',
        '利用端末': '',
        'メールキャリアコード': '',
        'あす楽希望フラグ': '',
        '医薬品受注フラグ': '',
        '楽天スーパーDEAL商品受注フラグ': '',
        'メンバーシッププログラム受注タイプ': '',
        '決済手数料': '',
        '注文者負担金合計': '',
        '店舗負担金合計': ''
      };
    });
  };  

  // バイナリ文字列をArrayBufferに変換
  const s2ab = (s: string): ArrayBuffer => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  };

  // Excelファイルとしてダウンロードする関数
  const downloadExcel = (data: ConvertedRow[]): void => {
    const ws = utils.json_to_sheet(data);
    const wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "変換データ");

    const fileName = "converted_data.xlsx";
    const wbout = write(wb, { bookType: 'xlsx', type: 'binary' });
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), fileName);
  };

  // ファイルアップロード処理
  const handleFileUpload = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    try {
      setLoading(true);
      setError(null);
      setSuccess(false);

      const file = event.target.files?.[0];
      if (!file) {
        throw new Error('ファイルが選択されていません。');
      }

      if (!file.name.endsWith('.csv')) {
        throw new Error('CSVファイルを選択してください。');
      }

      Papa.parse(file, {
        complete: (results: Papa.ParseResult<CSVRow>) => {
          try {
            if (results.errors.length > 0) {
              throw new Error('CSVファイルの解析中にエラーが発生しました。');
            }

            const converted = convertCSVData(results.data);
            setConvertedData(converted);
            downloadExcel(converted);
            setSuccess(true);
          } catch (err) {
            setError(err instanceof Error ? err.message : '不明なエラーが発生しました。');
          }
        },
        header: true,
        encoding: 'Shift-JIS',
        error: (error: Papa.ParseError) => {
          setError('ファイルの読み込み中にエラーが発生しました。');
        }
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : '不明なエラーが発生しました。');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="converter-container" >
      <p>クロスモール(.csv) → サチ(.xlsx)</p>
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
          <p>変換完了</p>
        </div>
      )}

      {/*<div className="instructions">
        <h2>使用方法</h2>
        <ol>
          <li>変換したいCSVファイルを選択してください。</li>
          <li>自動的に変換が開始されます。</li>
          <li>変換されたExcelファイルが自動的にダウンロードされます。</li>
        </ol>
      </div>*/}
    </div>
  );
};

export default CSVConverter;
