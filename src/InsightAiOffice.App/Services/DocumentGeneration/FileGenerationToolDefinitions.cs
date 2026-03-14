using System.Text.Json.Nodes;
using InsightCommon.AI;

namespace InsightAiOffice.App.Services.DocumentGeneration;

/// <summary>
/// AI ツール定義 — ファイル生成系
/// generate_report, generate_presentation, generate_spreadsheet
/// </summary>
public static class FileGenerationToolDefinitions
{
    public static ToolDefinition GenerateReport { get; } = new()
    {
        Name = "generate_report",
        Description = """
            Word (.docx) または HTML レポートを生成します。
            セクション種別: title, heading, summary, text, recommendation,
            bullet_list, table, comparison, chart, key_metrics, page_break。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "出力先ファイルパス（.docx or .html）。省略時はドキュメントフォルダに保存。"
                },
                "format": {
                    "type": "string",
                    "enum": ["docx", "html"],
                    "description": "出力形式（デフォルト: docx）"
                },
                "theme": {
                    "type": "string",
                    "enum": ["gold", "blue", "green", "red", "navy", "mono"],
                    "description": "カラーテーマ。gold=Ivory&Gold(デフォルト), blue=ブルー系, green=グリーン系, red=レッド系, navy=ダークネイビー, mono=モノクロ。ユーザーが色を指定した場合に使用。"
                },
                "title": {
                    "type": "string",
                    "description": "レポートタイトル"
                },
                "author": {
                    "type": "string",
                    "description": "作成者名"
                },
                "date": {
                    "type": "string",
                    "description": "作成日（YYYY-MM-DD）"
                },
                "sections": {
                    "type": "array",
                    "description": "レポートセクション配列",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": { "type": "string", "enum": ["title","heading","summary","text","recommendation","bullet_list","table","comparison","chart","key_metrics","page_break"] },
                            "title": { "type": "string" },
                            "level": { "type": "integer", "description": "見出しレベル (1-3)" },
                            "content": { "type": "string", "description": "本文テキスト" },
                            "items": { "type": "array", "items": { "type": "string" }, "description": "箇条書き項目" },
                            "tableData": {
                                "type": "object",
                                "properties": {
                                    "headers": { "type": "array", "items": { "type": "string" } },
                                    "rows": { "type": "array", "items": { "type": "array", "items": { "type": "string" } } }
                                }
                            },
                            "metrics": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "label": { "type": "string" },
                                        "value": { "type": "string" },
                                        "change": { "type": "string" },
                                        "trend": { "type": "string", "enum": ["positive","negative","neutral"] }
                                    }
                                }
                            }
                        },
                        "required": ["type"]
                    }
                }
            },
            "required": ["title", "sections"]
        }
        """),
    };

    public static ToolDefinition GeneratePresentation { get; } = new()
    {
        Name = "generate_presentation",
        Description = """
            PowerPoint (.pptx) プレゼンテーションを生成します。
            スライド種別: Title, Agenda, Data, Content。
            各スライドにタイトル、キーメッセージ、箇条書き、スピーカーノートを設定可能。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "出力先 .pptx ファイルパス。省略時はドキュメントフォルダに保存。"
                },
                "title": {
                    "type": "string",
                    "description": "プレゼンテーションタイトル"
                },
                "theme": {
                    "type": "string",
                    "enum": ["gold", "blue", "green", "red", "navy", "mono"],
                    "description": "カラーテーマ（デフォルト: gold）。ユーザーが色を指定した場合に使用。"
                },
                "slides": {
                    "type": "array",
                    "description": "スライド仕様の配列",
                    "items": {
                        "type": "object",
                        "properties": {
                            "order": { "type": "integer" },
                            "slideType": { "type": "string", "enum": ["Title","Agenda","Data","Content"] },
                            "title": { "type": "string" },
                            "keyMessage": { "type": "string" },
                            "bullets": { "type": "array", "items": { "type": "string" } },
                            "speakerNotes": { "type": "string" }
                        },
                        "required": ["order", "title"]
                    }
                }
            },
            "required": ["title", "slides"]
        }
        """),
    };

    public static ToolDefinition GenerateSpreadsheet { get; } = new()
    {
        Name = "generate_spreadsheet",
        Description = """
            Excel (.xlsx) スプレッドシートを生成します。
            複数シート対応。各シートにヘッダーとデータ行を設定可能。
            数値は自動検出してフォーマットされます。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "output_path": {
                    "type": "string",
                    "description": "出力先 .xlsx ファイルパス。省略時はドキュメントフォルダに保存。"
                },
                "title": {
                    "type": "string",
                    "description": "スプレッドシートタイトル"
                },
                "theme": {
                    "type": "string",
                    "enum": ["gold", "blue", "green", "red", "navy", "mono"],
                    "description": "カラーテーマ（デフォルト: gold）。ユーザーが色を指定した場合に使用。"
                },
                "sheets": {
                    "type": "array",
                    "description": "シートデータの配列",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": { "type": "string" },
                            "headers": { "type": "array", "items": { "type": "string" } },
                            "rows": { "type": "array", "items": { "type": "array", "items": { "type": "string" } } }
                        },
                        "required": ["name", "headers", "rows"]
                    }
                }
            },
            "required": ["title", "sheets"]
        }
        """),
    };

    public static ToolDefinition GeneratePresentationFromTemplate { get; } = new()
    {
        Name = "generate_presentation_from_template",
        Description = """
            既存の PowerPoint (.pptx) テンプレートにスライドを追加して新しい PPTX を生成します。
            テンプレートのデザイン・テーマを維持したまま、新しいスライドを末尾に追加します。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "テンプレート .pptx ファイルパス（添付ファイルから取得）"
                },
                "output_path": {
                    "type": "string",
                    "description": "出力先 .pptx ファイルパス。省略時はドキュメントフォルダに保存。"
                },
                "title": {
                    "type": "string",
                    "description": "プレゼンテーションタイトル"
                },
                "theme": {
                    "type": "string",
                    "enum": ["gold", "blue", "green", "red", "navy", "mono"],
                    "description": "カラーテーマ（デフォルト: gold）。ユーザーが色を指定した場合に使用。"
                },
                "slides": {
                    "type": "array",
                    "description": "追加するスライド仕様の配列",
                    "items": {
                        "type": "object",
                        "properties": {
                            "order": { "type": "integer" },
                            "slideType": { "type": "string", "enum": ["Title","Agenda","Data","Content"] },
                            "title": { "type": "string" },
                            "keyMessage": { "type": "string" },
                            "bullets": { "type": "array", "items": { "type": "string" } },
                            "speakerNotes": { "type": "string" }
                        },
                        "required": ["order", "title"]
                    }
                }
            },
            "required": ["template_path", "title", "slides"]
        }
        """),
    };

    public static ToolDefinition BatchGenerate { get; } = new()
    {
        Name = "batch_generate",
        Description = """
            差し込み印刷（メールマージ）: Word (.docx) テンプレートと Excel (.xlsx) データソースから
            複数の Word 文書を一括生成します。テンプレート内の {{列名}} プレースホルダーが
            Excel の各行データに置換されます。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "template_path": {
                    "type": "string",
                    "description": "Word テンプレート (.docx) ファイルパス。{{列名}} でプレースホルダーを定義。"
                },
                "data_source_path": {
                    "type": "string",
                    "description": "Excel データソース (.xlsx) ファイルパス。1行目がヘッダー。"
                },
                "output_directory": {
                    "type": "string",
                    "description": "出力先フォルダパス。省略時はドキュメントフォルダ。"
                },
                "filename_column": {
                    "type": "string",
                    "description": "出力ファイル名に使う Excel 列名。省略時は連番。"
                }
            },
            "required": ["template_path", "data_source_path"]
        }
        """),
    };

    /// <summary>添付 Word ファイルの内容を書き換えて新しいファイルとして保存</summary>
    public static ToolDefinition RewriteDocument { get; } = new()
    {
        Name = "rewrite_document",
        Description = """
            添付された Word (.docx) ファイルのフォーマット（レイアウト・スタイル・書式）を維持したまま、
            テキスト内容を書き換えて新しいファイルとして保存します。
            元ファイルのデザインやフォントはそのまま残ります。
            """,
        InputSchema = Parse("""
        {
            "type": "object",
            "properties": {
                "source_path": {
                    "type": "string",
                    "description": "書き換え元の .docx ファイルパス（添付ファイルのパス）"
                },
                "output_path": {
                    "type": "string",
                    "description": "出力先ファイルパス。省略時はプロジェクトフォルダに保存。"
                },
                "replacements": {
                    "type": "array",
                    "description": "テキスト置換のリスト",
                    "items": {
                        "type": "object",
                        "properties": {
                            "find": { "type": "string", "description": "検索テキスト" },
                            "replace": { "type": "string", "description": "置換テキスト" }
                        },
                        "required": ["find", "replace"]
                    }
                },
                "title": {
                    "type": "string",
                    "description": "成果物のタイトル（ファイル名にも使用）"
                }
            },
            "required": ["source_path", "replacements"]
        }
        """),
    };

    /// <summary>全ツール定義を取得</summary>
    public static List<ToolDefinition> GetAllTools() => new()
    {
        GenerateReport,
        GeneratePresentation,
        GenerateSpreadsheet,
        GeneratePresentationFromTemplate,
        RewriteDocument,
        BatchGenerate,
    };

    private static JsonObject Parse(string json) => JsonNode.Parse(json)!.AsObject();
}
