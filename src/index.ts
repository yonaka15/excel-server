import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { XlwingsRpcClient } from "./xlwings-rpc-client.js";

// xlwings-rpcクライアントの設定
const xlwingsClient = new XlwingsRpcClient(
  process.env.XLWINGS_HOST,
  process.env.XLWINGS_PORT ? parseInt(process.env.XLWINGS_PORT, 10) : undefined
);

// サーバーインスタンス作成
const server = new McpServer({
  name: "excel-server",
  version: "1.0.0",
});

// ------------- アプリケーション操作ツール -------------

server.tool(
  "excel_app_list",
  "すべての実行中のExcelアプリケーションを取得します",
  {},
  async () => {
    try {
      const result = await xlwingsClient.appList();
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_app_get",
  "指定されたPIDまたはアクティブなExcelアプリケーションを取得します",
  {
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ pid }) => {
    try {
      const result = await xlwingsClient.appGet(pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_app_create",
  "新しいExcelアプリケーションを作成します",
  {
    visible: z
      .boolean()
      .optional()
      .describe("表示するかどうか（デフォルト: true）"),
    addBook: z
      .boolean()
      .optional()
      .describe("ブックを追加するかどうか（デフォルト: true）"),
  },
  async ({ visible, addBook }) => {
    try {
      const result = await xlwingsClient.appCreate(visible, addBook);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_app_quit",
  "Excelアプリケーションを終了します",
  {
    pid: z.number().describe("アプリケーションのPID"),
    saveChanges: z
      .boolean()
      .optional()
      .describe("変更を保存するかどうか（デフォルト: true）"),
  },
  async ({ pid, saveChanges }) => {
    try {
      const result = await xlwingsClient.appQuit(pid, saveChanges);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ------------- ワークブック操作ツール -------------

server.tool(
  "excel_book_list",
  "すべての開いているワークブックを取得します",
  {
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ pid }) => {
    try {
      const result = await xlwingsClient.bookList(pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_book_get",
  "指定されたワークブックを取得します",
  {
    book: z.string().describe("ワークブック名"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, pid }) => {
    try {
      const result = await xlwingsClient.bookGet(book, pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_book_open",
  "ワークブックを開きます",
  {
    path: z.string().describe("ワークブックのパス"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
    readOnly: z
      .boolean()
      .optional()
      .describe("読み取り専用で開くかどうか（オプション）"),
    password: z.string().optional().describe("パスワード（オプション）"),
  },
  async ({ path, pid, readOnly, password }) => {
    try {
      const result = await xlwingsClient.bookOpen(
        path,
        pid,
        readOnly,
        password
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_book_create",
  "新しいワークブックを作成します",
  {
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ pid }) => {
    try {
      const result = await xlwingsClient.bookCreate(pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_book_close",
  "ワークブックを閉じます",
  {
    book: z.string().describe("ワークブック名"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
    save: z
      .boolean()
      .optional()
      .describe("保存するかどうか（オプション、デフォルト: true）"),
  },
  async ({ book, pid, save }) => {
    try {
      const result = await xlwingsClient.bookClose(book, pid, save);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ------------- シート操作ツール -------------

server.tool(
  "excel_sheet_list",
  "ワークブック内のすべてのシートを取得します",
  {
    book: z.string().describe("ワークブック名"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, pid }) => {
    try {
      const result = await xlwingsClient.sheetList(book, pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_sheet_get",
  "特定のシートを取得します",
  {
    book: z.string().describe("ワークブック名"),
    name: z.string().describe("シート名"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, name, pid }) => {
    try {
      const result = await xlwingsClient.sheetGet(book, name, pid);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ------------- ヘルスチェックツール -------------

server.tool(
  "excel_health_check",
  "xlwings-rpcサーバーの接続状態を確認します",
  {},
  async () => {
    try {
      const url = new URL(xlwingsClient["url"]);

      // 単純なGETリクエストを送信してサーバーの状態を確認
      const response = await fetch(
        `http://${url.hostname}:${url.port}/health`,
        {
          method: "GET",
        }
      ).catch((error) => {
        console.error("Fetch error:", error);
        throw error;
      });

      if (!response.ok) {
        return {
          content: [
            {
              type: "text",
              text: `xlwings-rpcサーバーの状態確認に失敗しました: HTTP ${response.status}`,
            },
          ],
          isError: true,
        };
      }

      const text = await response.text();

      return {
        content: [
          {
            type: "text",
            text: `xlwings-rpcサーバーの状態: OK\nレスポンス: ${text}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `xlwings-rpcサーバーへの接続に失敗しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ------------- レンジ操作ツール -------------

server.tool(
  "excel_range_get_value",
  "セル範囲の値を取得します",
  {
    book: z.string().describe("ワークブック名"),
    sheet: z.string().describe("シート名"),
    address: z.string().describe("セル範囲のアドレス（例: 'A1:B10'）"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, sheet, address, pid }) => {
    try {
      const result = await xlwingsClient.rangeGetValue(
        book,
        sheet,
        address,
        pid
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_range_set_value",
  "セル範囲に値を設定します",
  {
    book: z.string().describe("ワークブック名"),
    sheet: z.string().describe("シート名"),
    address: z.string().describe("セル範囲のアドレス（例: 'A1:B10'）"),
    value: z.any().describe("設定する値（単一の値または配列）"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, sheet, address, value, pid }) => {
    try {
      const result = await xlwingsClient.rangeSetValue(
        book,
        sheet,
        address,
        value,
        pid
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_range_get_formula",
  "セル範囲の数式を取得します",
  {
    book: z.string().describe("ワークブック名"),
    sheet: z.string().describe("シート名"),
    address: z.string().describe("セル範囲のアドレス（例: 'A1:B10'）"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, sheet, address, pid }) => {
    try {
      const result = await xlwingsClient.rangeGetFormula(
        book,
        sheet,
        address,
        pid
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

server.tool(
  "excel_range_set_formula",
  "セル範囲に数式を設定します",
  {
    book: z.string().describe("ワークブック名"),
    sheet: z.string().describe("シート名"),
    address: z.string().describe("セル範囲のアドレス（例: 'A1:B10'）"),
    formula: z.any().describe("設定する数式（単一の数式または配列）"),
    pid: z.number().optional().describe("アプリケーションのPID（オプション）"),
  },
  async ({ book, sheet, address, formula, pid }) => {
    try {
      const result = await xlwingsClient.rangeSetFormula(
        book,
        sheet,
        address,
        formula,
        pid
      );
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// ------------- 直接アクセスツール -------------

server.tool(
  "excel_direct",
  "xlwings-rpcサーバーに直接JSON-RPCリクエストを送信します",
  {
    method: z.string().describe("実行するxlwings-rpcメソッド名"),
    params: z
      .any()
      .optional()
      .describe("メソッドに渡すパラメータ（オプション）"),
  },
  async ({ method, params }) => {
    try {
      console.error(`Executing direct method: ${method} with params:`, params);
      const result = await xlwingsClient.request(method, params);
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `エラーが発生しました: ${
              error instanceof Error ? error.message : String(error)
            }`,
          },
        ],
        isError: true,
      };
    }
  }
);

// サーバー起動
async function main() {
  try {
    // StdioサーバーTransport
    const stdioTransport = new StdioServerTransport();
    await server.connect(stdioTransport);
    console.error("MCP Server running on stdio");
  } catch (error) {
    console.error("Fatal error:", error);
    process.exit(1);
  }
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
