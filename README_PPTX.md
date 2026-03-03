# PPTX 转图片工具

使用 `aspose.slides` 将 PowerPoint 文件的每一页转换为高质量图片。

## 特点

- ✅ 无需安装 LibreOffice 或 Microsoft PowerPoint
- ✅ 跨平台支持（Windows, macOS, Linux）
- ✅ 支持 .pptx 和 .ppt 格式
- ✅ 支持多种输出格式：PNG, JPEG, BMP, GIF, TIFF
- ✅ 2x 缩放输出，保证高清质量
- ✅ 使用 uv 管理 Python 环境

## 安装

确保已安装 [uv](https://github.com/astral-sh/uv)，然后运行：

```bash
uv sync
```

## 使用方法

### 基本用法

```bash
# 将 presentation.pptx 转换为 PNG 图片
uv run pptx_to_images.py presentation.pptx

# 指定输出目录
uv run pptx_to_images.py presentation.pptx ./my_output

# 指定图片格式（png/jpeg/bmp/gif/tiff）
uv run pptx_to_images.py presentation.pptx ./output jpeg
```

### 完整示例

```bash
# 转换文件，输出到默认目录（与文件名相同）
uv run pptx_to_images.py "FA12e_01 Notes.pptx"
# 输出: FA12e_01 Notes/slide_001.png, slide_002.png, ...

# 转换文件，输出到指定目录，使用 JPEG 格式
uv run pptx_to_images.py "FA12e_01 Notes.pptx" ./slides jpg
```

## 输出结构

```
presentation.pptx
├── presentation/           # 默认输出目录（与PPT文件名相同）
│   ├── slide_001.png
│   ├── slide_002.png
│   ├── slide_003.png
│   └── ...
```

## 关于 aspose.slides

此工具使用 [Aspose.Slides for Python](https://products.aspose.com/slides/python-net/)，这是一个商业级库。

- **评估模式**: 免费使用，但输出图片会有水印
- **许可证**: 如需去除水印，需要购买许可证

## 替代方案

如果不想使用 aspose.slides，可以考虑以下替代方案：

1. **LibreOffice + pptx2img** (免费，需要安装 LibreOffice)
   ```bash
   uv add pptx2img
   ```

2. **python-pptx + 截图工具** (仅支持提取已有图片)

3. **Windows COM + pywin32** (仅 Windows，需要安装 PowerPoint)
