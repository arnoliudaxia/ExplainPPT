# ExplainPPT

将 PowerPoint 文件的幻灯片转换为 3x3 九宫格拼图。

## 特点

- ✅ 无需安装 LibreOffice 或 Microsoft PowerPoint
- ✅ 跨平台支持（Windows, macOS, Linux）
- ✅ 支持 .pptx 和 .ppt 格式
- ✅ 使用 CloudConvert API 进行云端转换
- ✅ 自动将幻灯片按 9 张一组拼接成九宫格
- ✅ 使用 uv 管理 Python 环境

## 安装

确保已安装 [uv](https://github.com/astral-sh/uv)，然后运行：

```bash
uv sync
```

## 配置

### 方法 1: 使用 .env 文件（推荐）

复制示例文件并填写你的 API Key：

```bash
cp .env.example .env
# 编辑 .env 文件，填入你的 API Key
```

`.env` 文件内容：
```
CLOUDCONVERT_API_KEY=your_api_key_here
```

### 方法 2: 使用环境变量

```bash
# Linux/macOS
export CLOUDCONVERT_API_KEY=your_api_key

# Windows
set CLOUDCONVERT_API_KEY=your_api_key
```

获取 API Key: https://cloudconvert.com/dashboard/api

## 使用方法

### 基本用法

```bash
# 将 presentation.pptx 转换为九宫格
uv run ppt_to_grid.py presentation.pptx

# 指定输出目录
uv run ppt_to_grid.py presentation.pptx ./my_output
```

### 完整示例

```bash
# 转换文件，输出到默认目录（文件名_output）
uv run ppt_to_grid.py "FA12e_01 Notes.pptx"

# 转换文件，输出到指定目录
uv run ppt_to_grid.py "FA12e_01 Notes.pptx" ./grids
```

## 输出结构

```
presentation.pptx
├── presentation_output/        # 默认输出目录
│   ├── slides/                 # 转换后的单页幻灯片图片
│   │   ├── slide-1.jpg
│   │   ├── slide-2.jpg
│   │   └── ...
│   └── grids/                  # 九宫格拼图
│       ├── grid_01.jpg         # 幻灯片 1-9
│       ├── grid_02.jpg         # 幻灯片 10-18
│       └── ...
```

## 工具脚本

### grid_collage.py

独立的九宫格拼图工具，用于将已有图片拼接成 3x3 网格：

```bash
# 将当前目录的所有 jpg 图片按 9 张一组拼接
cd images_directory
uv run grid_collage.py
```

## 依赖

- **CloudConvert**: 用于 PPT 转图片（需要 API Key）
- **Pillow**: 用于图片处理和九宫格拼接

## 替代方案

如果不想使用 CloudConvert API，可以考虑以下替代方案：

1. **LibreOffice + 命令行转换** (免费，需要安装 LibreOffice)
   ```bash
   libreoffice --headless --convert-to pdf presentation.pptx
   pdftoppm -jpeg presentation.pdf slide
   ```

2. **Windows COM + pywin32** (仅 Windows，需要安装 PowerPoint)

## 许可证

MIT License
