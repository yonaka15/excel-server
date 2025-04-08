import { randomUUID } from "crypto";

/**
 * xlwings-rpc APIクライアント
 * JSON-RPC 2.0プロトコルを使用してxlwings-rpcサーバーにリクエストを送信します
 */
export class XlwingsRpcClient {
  private url: string;
  
  /**
   * コンストラクタ
   * @param host ホスト（デフォルト: '0.0.0.0'）
   * @param port ポート（デフォルト: 8000）
   */
  constructor(host?: string, port?: number) {
    const serverHost = host || process.env.XLWINGS_HOST || '0.0.0.0';
    const serverPort = port || parseInt(process.env.XLWINGS_PORT || '8000', 10);
    this.url = `http://${serverHost}:${serverPort}/rpc`;
  }

  /**
   * Python形式の値をJavaScriptオブジェクトに変換する
   * @param text Python形式の文字列
   * @returns パース済みのJavaScriptオブジェクト
   */
  private parsePythonLikeJson(text: string): any {
    // Python の True/False を JavaScript の true/false に置換
    const processedText = text
      .replace(/True/g, 'true')
      .replace(/False/g, 'false')
      .replace(/None/g, 'null');
    
    // VersionNumber("x.y") のような表現を "x.y" に置換
    const versionNumberRegex = /VersionNumber\(["']([^"']*)["']\)/g;
    const processedText2 = processedText.replace(versionNumberRegex, '"$1"');

    // 他のPythonカスタムクラスを処理
    const pythonClassRegex = /([a-zA-Z]+)\(([^)]*)\)/g;
    const processedText3 = processedText2.replace(pythonClassRegex, (match, className, args) => {
      // クラス名と引数を取得
      try {
        // JSON文字列として解析可能かチェック
        const parsedArgs = JSON.parse(`[${args}]`);
        return JSON.stringify({ className, args: parsedArgs });
      } catch (e) {
        // 解析できない場合は元の文字列をJSON文字列として返す
        return `"${match}"`;
      }
    });
    
    try {
      return JSON.parse(processedText3);
    } catch (error) {
      console.error("Failed to parse processed text:", processedText3);
      throw new Error(`Failed to parse Python-like JSON: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * JSON-RPCリクエストを送信します
   * @param method メソッド名
   * @param params パラメータ（オプション）
   * @returns レスポンス結果
   */
  async request<T = any>(method: string, params?: any): Promise<T> {
    const requestId = randomUUID();
    const requestBody = {
      jsonrpc: "2.0",
      method,
      params,
      id: requestId
    };

    try {
      console.log(`Sending request to ${this.url}:`, JSON.stringify(requestBody));
      
      const response = await fetch(this.url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // レスポンスのテキストを取得
      const responseText = await response.text();
      
      // レスポンスの内容をログ出力（デバッグ用）
      console.log("Response from xlwings-rpc server:", responseText);
      
      let data;
      try {
        // Python形式のJSONを処理
        data = this.parsePythonLikeJson(responseText);
      } catch (parseError) {
        console.error("Parse error:", parseError);
        throw new Error(`Invalid response format from server: ${responseText.substring(0, 100)}...`);
      }

      if (data.error) {
        throw new XlwingsRpcError(data.error.code, data.error.message, data.error.data);
      }

      return data.result as T;
    } catch (error) {
      if (error instanceof XlwingsRpcError) {
        throw error;
      }
      throw new Error(`Request failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * バッチリクエストを送信します
   * @param requests リクエストの配列
   * @returns レスポンスの配列
   */
  async batchRequest(requests: { method: string; params?: any }[]): Promise<any[]> {
    const requestBody = requests.map((req, index) => ({
      jsonrpc: "2.0",
      method: req.method,
      params: req.params,
      id: index + 1
    }));

    try {
      console.log(`Sending batch request to ${this.url}:`, JSON.stringify(requestBody));
      
      const response = await fetch(this.url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // レスポンスのテキストを取得
      const responseText = await response.text();
      
      // レスポンスの内容をログ出力（デバッグ用）
      console.log("Batch response from xlwings-rpc server:", responseText);
      
      let data;
      try {
        // Python形式のJSONを処理
        data = this.parsePythonLikeJson(responseText);
      } catch (parseError) {
        console.error("Parse error:", parseError);
        throw new Error(`Invalid response format from server: ${responseText.substring(0, 100)}...`);
      }

      // エラーチェック
      for (const item of data) {
        if (item.error) {
          throw new XlwingsRpcError(item.error.code, item.error.message, item.error.data);
        }
      }

      // レスポンスを元のリクエスト順に並び替え
      return data
        .sort((a: any, b: any) => a.id - b.id)
        .map((item: any) => item.result);
    } catch (error) {
      if (error instanceof XlwingsRpcError) {
        throw error;
      }
      throw new Error(`Batch request failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  // ------------- アプリケーション操作 -------------

  /**
   * すべての実行中のExcelアプリケーションを取得します
   */
  async appList() {
    return this.request('app.list');
  }

  /**
   * 指定されたPIDまたはアクティブなExcelアプリケーションを取得します
   * @param pid アプリケーションのPID（オプション）
   */
  async appGet(pid?: number) {
    return this.request('app.get', pid ? { pid } : undefined);
  }

  /**
   * 新しいExcelアプリケーションを作成します
   * @param visible 表示するかどうか（デフォルト: true）
   * @param addBook ブックを追加するかどうか（デフォルト: true）
   */
  async appCreate(visible?: boolean, addBook?: boolean) {
    const params: any = {};
    if (visible !== undefined) params.visible = visible;
    if (addBook !== undefined) params.add_book = addBook;
    return this.request('app.create', params);
  }

  /**
   * Excelアプリケーションを終了します
   * @param pid アプリケーションのPID
   * @param saveChanges 変更を保存するかどうか（デフォルト: true）
   */
  async appQuit(pid: number, saveChanges?: boolean) {
    const params: any = { pid };
    if (saveChanges !== undefined) params.save_changes = saveChanges;
    return this.request('app.quit', params);
  }

  /**
   * 計算モードを設定します
   * @param pid アプリケーションのPID
   * @param mode 計算モード（'automatic'|'manual'|'semiautomatic'）
   */
  async appSetCalculation(pid: number, mode: 'automatic' | 'manual' | 'semiautomatic') {
    return this.request('app.set_calculation', { pid, mode });
  }

  /**
   * 現在の計算モードを取得します
   * @param pid アプリケーションのPID
   */
  async appGetCalculation(pid: number) {
    return this.request('app.get_calculation', { pid });
  }

  /**
   * 指定されたアプリケーションで開いているワークブックを取得します
   * @param pid アプリケーションのPID
   */
  async appGetBooks(pid: number) {
    return this.request('app.get_books', { pid });
  }

  // ------------- ワークブック操作 -------------

  /**
   * すべての開いているワークブックを取得します
   * @param pid アプリケーションのPID（オプション）
   */
  async bookList(pid?: number) {
    return this.request('book.list', pid ? { pid } : undefined);
  }

  /**
   * 指定されたワークブックを取得します
   * @param name ワークブック名
   * @param pid アプリケーションのPID（オプション）
   */
  async bookGet(name: string, pid?: number) {
    const params: any = { name };
    if (pid !== undefined) params.pid = pid;
    return this.request('book.get', params);
  }

  /**
   * ワークブックを開きます
   * @param path ワークブックのパス
   * @param pid アプリケーションのPID（オプション）
   * @param readOnly 読み取り専用で開くかどうか（オプション）
   * @param password パスワード（オプション）
   */
  async bookOpen(path: string, pid?: number, readOnly?: boolean, password?: string) {
    const params: any = { path };
    if (pid !== undefined) params.pid = pid;
    if (readOnly !== undefined) params.read_only = readOnly;
    if (password !== undefined) params.password = password;
    return this.request('book.open', params);
  }

  /**
   * 新しいワークブックを作成します
   * @param pid アプリケーションのPID（オプション）
   */
  async bookCreate(pid?: number) {
    return this.request('book.create', pid ? { pid } : undefined);
  }

  /**
   * ワークブックを閉じます
   * @param name ワークブック名
   * @param pid アプリケーションのPID（オプション）
   * @param save 保存するかどうか（オプション、デフォルト: true）
   */
  async bookClose(name: string, pid?: number, save?: boolean) {
    const params: any = { name };
    if (pid !== undefined) params.pid = pid;
    if (save !== undefined) params.save = save;
    return this.request('book.close', params);
  }

  /**
   * ワークブックを保存します
   * @param name ワークブック名
   * @param pid アプリケーションのPID（オプション）
   * @param path 保存先パス（オプション）
   */
  async bookSave(name: string, pid?: number, path?: string) {
    const params: any = { name };
    if (pid !== undefined) params.pid = pid;
    if (path !== undefined) params.path = path;
    return this.request('book.save', params);
  }

  /**
   * ワークブック内のすべてのシートを取得します
   * @param name ワークブック名
   * @param pid アプリケーションのPID（オプション）
   */
  async bookGetSheets(name: string, pid?: number) {
    const params: any = { name };
    if (pid !== undefined) params.pid = pid;
    return this.request('book.get_sheets', params);
  }

  // ------------- シート操作 -------------

  /**
   * ワークブック内のすべてのシートを取得します
   * @param book ワークブック名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetList(book: string, pid?: number) {
    const params: any = { book };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.list', params);
  }

  /**
   * 特定のシートを取得します
   * @param book ワークブック名
   * @param name シート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetGet(book: string, name: string, pid?: number) {
    const params: any = { book, name };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.get', params);
  }

  /**
   * 新しいシートを追加します
   * @param book ワークブック名
   * @param name シート名（オプション）
   * @param before 既存のシート名（オプション）
   * @param after 既存のシート名（オプション）
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetAdd(book: string, name?: string, before?: string, after?: string, pid?: number) {
    const params: any = { book };
    if (name !== undefined) params.name = name;
    if (before !== undefined) params.before = before;
    if (after !== undefined) params.after = after;
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.add', params);
  }

  /**
   * シートを削除します
   * @param book ワークブック名
   * @param name シート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetDelete(book: string, name: string, pid?: number) {
    const params: any = { book, name };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.delete', params);
  }

  /**
   * シートの名前を変更します
   * @param book ワークブック名
   * @param name 現在のシート名
   * @param newName 新しいシート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetRename(book: string, name: string, newName: string, pid?: number) {
    const params: any = { book, name, new_name: newName };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.rename', params);
  }

  /**
   * シートの内容をクリアします
   * @param book ワークブック名
   * @param name シート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetClear(book: string, name: string, pid?: number) {
    const params: any = { book, name };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.clear', params);
  }

  /**
   * シートの使用範囲を取得します
   * @param book ワークブック名
   * @param name シート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetGetUsedRange(book: string, name: string, pid?: number) {
    const params: any = { book, name };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.get_used_range', params);
  }

  /**
   * シートをアクティブにします
   * @param book ワークブック名
   * @param name シート名
   * @param pid アプリケーションのPID（オプション）
   */
  async sheetActivate(book: string, name: string, pid?: number) {
    const params: any = { book, name };
    if (pid !== undefined) params.pid = pid;
    return this.request('sheet.activate', params);
  }

  // ------------- レンジ操作 -------------

  /**
   * 特定のセル範囲を取得します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeGet(book: string, sheet: string, address: string, pid?: number) {
    const params: any = { book, sheet, address };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.get', params);
  }

  /**
   * セル範囲の値を取得します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeGetValue(book: string, sheet: string, address: string, pid?: number) {
    const params: any = { book, sheet, address };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.get_value', params);
  }

  /**
   * セル範囲に値を設定します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param value 設定する値
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeSetValue(book: string, sheet: string, address: string, value: any, pid?: number) {
    const params: any = { book, sheet, address, value };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.set_value', params);
  }

  /**
   * セル範囲の数式を取得します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeGetFormula(book: string, sheet: string, address: string, pid?: number) {
    const params: any = { book, sheet, address };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.get_formula', params);
  }

  /**
   * セル範囲に数式を設定します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param formula 設定する数式
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeSetFormula(book: string, sheet: string, address: string, formula: any, pid?: number) {
    const params: any = { book, sheet, address, formula };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.set_formula', params);
  }

  /**
   * セル範囲をクリアします
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeClear(book: string, sheet: string, address: string, pid?: number) {
    const params: any = { book, sheet, address };
    if (pid !== undefined) params.pid = pid;
    return this.request('range.clear', params);
  }

  /**
   * セル範囲をpandas DataFrameとして取得します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param header ヘッダーを含むか（オプション、デフォルト: true）
   * @param index インデックスを含むか（オプション、デフォルト: false）
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeGetAsDataframe(
    book: string,
    sheet: string,
    address: string,
    header?: boolean,
    index?: boolean,
    pid?: number
  ) {
    const params: any = { book, sheet, address };
    if (header !== undefined) params.header = header;
    if (index !== undefined) params.index = index;
    if (pid !== undefined) params.pid = pid;
    return this.request('range.get_as_dataframe', params);
  }

  /**
   * pandas DataFrameをセル範囲に設定します
   * @param book ワークブック名
   * @param sheet シート名
   * @param address セル範囲のアドレス
   * @param dataframe DataFrameオブジェクト
   * @param header ヘッダーを含むか（オプション、デフォルト: true）
   * @param index インデックスを含むか（オプション、デフォルト: false）
   * @param pid アプリケーションのPID（オプション）
   */
  async rangeSetDataframe(
    book: string,
    sheet: string,
    address: string,
    dataframe: any,
    header?: boolean,
    index?: boolean,
    pid?: number
  ) {
    const params: any = { book, sheet, address, dataframe };
    if (header !== undefined) params.header = header;
    if (index !== undefined) params.index = index;
    if (pid !== undefined) params.pid = pid;
    return this.request('range.set_dataframe', params);
  }
}

/**
 * xlwings-rpc APIエラークラス
 */
export class XlwingsRpcError extends Error {
  constructor(
    public code: number,
    message: string,
    public data?: any
  ) {
    super(message);
    this.name = 'XlwingsRpcError';
  }
}
