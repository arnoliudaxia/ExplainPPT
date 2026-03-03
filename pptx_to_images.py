#!/usr/bin/env python3
"""
PPTX 转图片工具
使用 aspose.slides 将 PowerPoint 文件的每一页转换为图片

安装依赖:
    uv add aspose.slides

使用方法:
    uv run pptx_to_images.py <pptx文件路径> [输出目录] [图片格式]

示例:
    uv run pptx_to_images.py presentation.pptx
    uv run pptx_to_images.py presentation.pptx ./output png
"""

import sys
import os
from pathlib import Path

def convert_pptx_to_images(pptx_path, output_dir=None, image_format="png"):
    """
    将 PPTX 文件的每一页转换为图片
    
    Args:
        pptx_path: PPTX 文件路径
        output_dir: 输出目录，默认为 PPTX 文件名（不含扩展名）
        image_format: 图片格式，支持 png, jpeg, bmp, gif, tiff
    """
    try:
        import aspose.slides as slides
    except ImportError:
        print("错误: 未安装 aspose.slides")
        print("请运行: uv add aspose.slides")
        sys.exit(1)
    
    pptx_path = Path(pptx_path)
    
    if not pptx_path.exists():
        print(f"错误: 文件不存在: {pptx_path}")
        sys.exit(1)
    
    if not pptx_path.suffix.lower() in ['.pptx', '.ppt']:
        print(f"错误: 不支持的文件格式: {pptx_path.suffix}")
        print("只支持 .pptx 和 .ppt 格式")
        sys.exit(1)
    
    # 设置输出目录
    if output_dir is None:
        output_dir = pptx_path.stem
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # 验证图片格式
    valid_formats = ['png', 'jpeg', 'jpg', 'bmp', 'gif', 'tiff']
    image_format = image_format.lower().replace('jpg', 'jpeg')
    if image_format not in valid_formats:
        print(f"错误: 不支持的图片格式: {image_format}")
        print(f"支持的格式: {', '.join(valid_formats)}")
        sys.exit(1)
    
    print(f"正在加载: {pptx_path}")
    
    # 加载演示文稿
    presentation = slides.Presentation(str(pptx_path))
    
    # 获取格式对应的 ImageFormat
    format_map = {
        'png': slides.ImageFormat.PNG,
        'jpeg': slides.ImageFormat.JPEG,
        'bmp': slides.ImageFormat.BMP,
        'gif': slides.ImageFormat.GIF,
        'tiff': slides.ImageFormat.TIFF,
    }
    
    slide_count = len(presentation.slides)
    print(f"共 {slide_count} 页幻灯片")
    print(f"输出目录: {output_path.absolute()}")
    print(f"图片格式: {image_format.upper()}")
    print("-" * 50)
    
    # 转换每一页
    for i, slide in enumerate(presentation.slides, 1):
        # 生成文件名
        filename = f"slide_{i:03d}.{image_format.replace('jpeg', 'jpg')}"
        filepath = output_path / filename
        
        # 导出图片
        # 使用 2x 缩放以获得更高质量
        slide.get_image(2.0, 2.0).save(str(filepath), format_map[image_format])
        
        print(f"[OK] 第 {i}/{slide_count} 页 -> {filename}")
    
    print("-" * 50)
    print(f"完成！共导出 {slide_count} 张图片到: {output_path.absolute()}")
    
    return output_path


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n参数:")
        print("  <pptx文件路径>  要转换的 PowerPoint 文件")
        print("  [输出目录]      可选，默认为文件名（不含扩展名）")
        print("  [图片格式]      可选，默认为 png，支持 png/jpeg/bmp/gif/tiff")
        sys.exit(1)
    
    pptx_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    image_format = sys.argv[3] if len(sys.argv) > 3 else "png"
    
    convert_pptx_to_images(pptx_file, output_dir, image_format)


if __name__ == "__main__":
    main()
