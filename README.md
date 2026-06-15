# 案件线索移送函批量生成工具

Excel 解析 + Word 移送函批量生成 + 台账生成的桌面应用，支持 macOS 和 Windows。

## 功能

- 解析工单 Excel，自动分类（外卖/三方公司/三方个人/自营开票）
- 批量生成对应的 Word 移送函
- 生成台账 Excel
- 登记机关查询（mock 地址启发式 / 天眼查 API / 企查查 API）
- 桌面 GUI 操作界面

## 直接运行（开发模式）

需要 Python 3.10+，安装依赖后运行：

```bash
pip install flask pywebview openpyxl python-docx requests
python web.py
```

## 打包为桌面应用

### macOS

```bash
pip install pyinstaller
python build.py --mac
```

输出：`dist/移送函生成工具.app`（可直接双击运行）

### Windows

在 Windows 机器上执行：

```powershell
pip install pyinstaller
python build.py --win
```

输出：`dist/移送函生成工具.exe`（单文件可执行程序，无需 Python 环境）

### 清理构建产物

```bash
python build.py --clean
```

## 项目结构

```
lsy/
├── web.py              # 桌面 GUI 入口 (Flask + pywebview)
├── main.py             # CLI 入口 (保留)
├── config.py           # 配置（路径、模板、API）
├── excel_parser.py     # Excel 解析 + 分类
├── word_generator.py   # Word 模板填充生成
├── ledger_generator.py # 台账 Excel 生成
├── api_client.py       # 登记机关查询（可插拔 API）
├── build.py            # 跨平台构建脚本
├── build.spec          # PyInstaller spec (可选)
├── demands/            # 源数据 + 模板文件
└── output/             # 生成结果输出目录
```

## API 配置

默认使用 `mock` 模式（地址启发式匹配登记机关），覆盖大部分国内地址。

如需接入真实 API，在 GUI 中选择对应提供商并填写 API Key：

- **天眼查**: 需 API Token（天眼查开放平台申请）
- **企查查**: 需 API Key + Secret Key（企查查开放平台申请）

程序优先调 API 查询，失败时自动回退到地址启发式匹配。

## 跨平台说明

| 功能         | macOS                | Windows              |
|-------------|----------------------|----------------------|
| 打开输出目录  | Finder (`open`)      | Explorer (`startfile`) |
| 路径分隔符   | `/`                  | `\`                  |
| 打包格式     | `.app` bundle        | `.exe` 单文件         |
| PyInstaller  | `--windowed`         | `--windowed`         |
