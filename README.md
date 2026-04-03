# SatsueiSlip

ムービー納品時に同梱する「納品伝票」をPDF出力する、Windows向けデスクトップアプリです。

ドラッグ＆ドロップした動画ファイル、または動画フォルダーを `ffprobe` で解析し、一覧確認と備考編集をしてから納品伝票PDFを書き出せます。

## 主な機能

- 会社名、作品名、納品日、宛先、担当者名、備考の入力
- `.mp4` `.mov` `.avi` `.mxf` `.mkv` のドラッグ＆ドロップ読み込み
- フォルダー投入時のサブフォルダー再帰探索
- ffprobe による解像度、fps、総フレーム数、再生時間、ファイルサイズ取得
- 総フレーム数が直接取れない場合の `duration * fps` 推定
- 一覧表での備考編集、選択行削除、一覧クリア、再読み込み
- A4縦の日本語PDF出力、複数ページ対応
- 会社名、作品名、前回PDF保存先、ウィンドウサイズの復元

## セットアップ手順

### 1. Python を用意

Python 3.10 以上をインストールしてください。

### 2. 依存ライブラリをインストール

```powershell
cd D:\Documents\GitHub\SatsueiSlip
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
```

### 3. ffprobe を用意

どちらかの方法で `ffprobe.exe` を使える状態にしてください。

#### 方法A: PATH に ffprobe を通す

FFmpeg公式ビルドなどを導入し、`ffprobe.exe` を `PATH` に追加してください。

確認:

```powershell
ffprobe -version
```

#### 方法B: アプリ配下に配置

以下のどちらかに `ffprobe.exe` を配置してください。

- `tools\ffprobe\ffprobe.exe`
- `tools\ffprobe\bin\ffprobe.exe`

アプリは起動時に上記パスと `PATH` を順に探します。

## 開発環境での起動手順

```powershell
cd D:\Documents\GitHub\SatsueiSlip
.\.venv\Scripts\Activate.ps1
python -m satsuei_slip
```

または:

```powershell
satsuei-slip
```

## ビルド手順

PyInstaller で exe 化できます。

```powershell
cd D:\Documents\GitHub\SatsueiSlip
.\.venv\Scripts\Activate.ps1
python -m pip install pyinstaller
pyinstaller .\SatsueiSlip.spec
```

生成物:

- `dist\SatsueiSlip\SatsueiSlip.exe`

`ffprobe.exe` を同梱したい場合は、`tools\ffprobe\ffprobe.exe` または `tools\ffprobe\bin\ffprobe.exe` に配置してからビルドしてください。

## インストーラー作成手順

Inno Setup 6 を使って Windows インストーラーを作成します。

### 1. Inno Setup 6 をインストール

`ISCC.exe` が使えるように、Inno Setup 6 をインストールしてください。

### 2. インストーラーをビルド

```powershell
cd D:\Documents\GitHub\SatsueiSlip
.\.venv\Scripts\Activate.ps1
.\scripts\build_installer.ps1
```

生成物:

- `installer_output\SatsueiSlip_Setup_0.1.0.exe`

インストーラー定義は `installer\SatsueiSlip.iss` です。

## GitHub から更新を取得する設定

アプリの `ヘルプ` → `GitHubの更新を確認` から、GitHub Releases の最新版を確認できます。

### 1. リポジトリ名を設定

`src\satsuei_slip\release_config.py` の `GITHUB_OWNER` を、自分の GitHub ユーザー名または Organization 名に変更してください。

```python
GITHUB_OWNER = "your-github-user"
GITHUB_REPO = "SatsueiSlip"
```

### 2. GitHub Releases に新バージョンを公開

1. `src\satsuei_slip\__init__.py` の `__version__` を更新します。
2. `installer\SatsueiSlip.iss` の `MyAppVersion` を同じバージョンへ更新します。
3. `.\scripts\build_installer.ps1` で新しいインストーラーを作成します。
4. GitHub の Releases で、タグ名を `v0.1.1` のように付けてインストーラーexeを添付します。

アプリは `https://api.github.com/repos/<owner>/<repo>/releases/latest` を参照し、現在の `__version__` より新しいタグがあればリリースページを開けるようにします。

## 操作手順

1. 会社名、作品名などの基本情報を入力します。
2. 中央のドロップエリアへ動画ファイル、または動画フォルダーをドロップします。
3. 一覧で解析結果を確認し、必要なら備考を編集します。
4. `PDF書き出し` を押して保存先を選びます。

## 注意

- 会社名、作品名、動画1件以上が揃っていない場合はPDF出力できません。
- フレーム数が直接取得できない動画は推定値として扱い、PDF上では `*` を付けて表示します。
- 可変フレームレート素材ではフレーム数が推定になる場合があります。
- 日本語ファイル名、全角パスを想定して `pathlib.Path` と `subprocess` の引数配列で処理しています。
