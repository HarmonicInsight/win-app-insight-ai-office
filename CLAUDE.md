# CLAUDE.md - Project Notes

## Project Overview

- **App Name**: Insight AI Office
- **Publisher**: HARMONIC insight
- **Framework**: .NET 8.0 (WPF, WinExe)
- **Installer**: Inno Setup 6.x (`Installer/InsightAiOffice.iss`)
- **Solution**: `InsightAiOffice.sln`

## App Icon Setup

アプリアイコンは `src/InsightAiOffice.App/Resources/InsightAiOffice.ico` に配置。

### 適用箇所

1. **アプリケーション本体** — `.csproj` の `<ApplicationIcon>` で設定済み
2. **全ウィンドウ** — 各 XAML の `Icon` 属性で設定済み
3. **インストーラー** — `InsightAiOffice.iss` の `SetupIconFile` でインストーラー自体のアイコンに使用。スタートメニュー・デスクトップショートカットにも `IconFilename` で指定済み

### アイコン生成元

- 元画像: `src/InsightAiOffice.App/Resources/icon.png`
- ICO変換: ImageMagick で 256x256, 128x128, 64x64, 48x48, 32x32, 16x16 の各サイズを含むマルチサイズICOとして生成

## Build & Publish

```bash
# Publish (single-file, self-contained)
dotnet publish src/InsightAiOffice.App -c Release -r win-x64 --self-contained -o publish

# Build installer (requires Inno Setup)
ISCC.exe Installer/InsightAiOffice.iss
```

## Directory Structure (Key Paths)

- `src/InsightAiOffice.App/` — WPF アプリ本体
- `src/InsightAiOffice.Core/` — コアロジック
- `src/InsightAiOffice.Data/` — データ層
- `tests/InsightAiOffice.Core.Tests/` — テスト
- `Installer/` — Inno Setup スクリプト
- `Output/` — インストーラー出力先 (.gitignore)
- `publish/` — dotnet publish 出力先 (.gitignore)
