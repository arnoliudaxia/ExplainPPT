#!/usr/bin/env python3
"""
PPT 转九宫格拼图工具

工作流程:
1. 使用 CloudConvert 将 PPT 转换为图片
2. 使用 grid_collage 将图片按顺序拼成 3x3 九宫格

环境变量:
    CLOUDCONVERT_API_KEY: CloudConvert API Key

使用方法:
    # 设置 API Key (Windows)
    set CLOUDCONVERT_API_KEY=your_api_key
    
    # 运行脚本
    uv run ppt_to_grid.py <pptx文件路径> [输出目录]

示例:
    uv run ppt_to_grid.py "2/FA12e_02_Notes.pptx"
    uv run ppt_to_grid.py "presentation.pptx" ./output
"""

import os
import sys
import subprocess
import time
import glob
import re
from pathlib import Path
from PIL import Image


def natural_sort_key(s):
    """自然排序，正确处理数字"""
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', str(s))]


def convert_pptx_with_cloudconvert(pptx_path, output_dir, image_format="jpg"):
    """
    使用 CloudConvert CLI 将 PPT 转换为图片
    
    Args:
        pptx_path: PPTX 文件路径
        output_dir: 输出目录
        image_format: 图片格式 (jpg 或 png)
    
    Returns:
        生成的图片路径列表
    """
    pptx_path = Path(pptx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 检查 API Key（支持环境变量或直接传入）
    api_key = os.environ.get('CLOUDCONVERT_API_KEY')
    if not api_key:
        # 尝试使用硬编码的 API Key（仅用于测试）
        api_key = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiY2JmMmUwZTNkYTc0MWNlMmI1N2I4Y2ZhYmQzNGM5OGRlNmJlYzFlMDU3YTQ3OWJiNzRmYTBjNjYyYTc0N2M1NmMwOTA5ZDA4NDg5YmNhZjkiLCJpYXQiOjE3NzI0NDk0OTAuMjM3MDA5LCJuYmYiOjE3NzI0NDk0OTAuMjM3MDEsImV4cCI6NDkyODEyMzA5MC4yMzIyODcsInN1YiI6Ijc0NTI1Nzg2Iiwic2NvcGVzIjpbInVzZXIucmVhZCIsInRhc2sucmVhZCIsInRhc2sud3JpdGUiXX0.RcL52fSqQnhynBBTlNesH1pu1MPQ1ERaANyvF41Hcs5O_SKOKapmvN4cpQr8qv1vpwTCuMQtX0b_7nOrEimvDCdT2g-4zuGdR64-yjRnB6loo_F80y0Eo_5kWdOB3zvoMK2qZwab3LqitOt4wZfb_zAnvVvJtvqUBAMNwlmmwNNlUb7Tx4QcENrVgQQXC9VDEspSJ9M-lUyfnzmuN1SYHRzOJPKlZcl2lSM_jCNC6xvJUnciOaV_aXhx8ZN1d0Z1gnk3jzv5vbAbCPAeRnjUKkQB9p9HLibaT2DkWDZA9jFWrLO0dBXYTjoaE2Kw6yY0AkABeA4Up_r1U3lM0t21QC-UwQmUrWTKbJhogO1rZcFYzKXmdMsTD-yQGv7QDgEt31UzDtv6OHRLNVFWw-rTJA7LyIcTbs33CNnRjfjMLuo0duLUFgfA_9Btr6Eb7KmXJe_i_qnfNPfuKNOz01mYRfa9dgJ6Cest0yTL8y48ewpqHwMLjUitrww-0qAU0twMVtKVCr4TSWr9aJJKSHlC42ItDLWFIfJ6lmNEu90WGEJNwj_uCEqLB5UhMvCys-Qf2wf9TAs0It5b5Km-ClvvvKhv7k0RcijenZ_zbWkvDwgM7b_pHQVoMsiUISxuXoEZNE9x6f-2Hs47IgGoTIXrJ6FwKCwy1rI-KTMRXgNgWY4"
        os.environ['CLOUDCONVERT_API_KEY'] = api_key
    
    # 检查 cloudconvert CLI
    try:
        result = subprocess.run(['cloudconvert', '--version'], 
                              capture_output=True, text=True, check=True)
        print(f"CloudConvert CLI: {result.stdout.strip()}")
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("错误: 未安装 CloudConvert CLI")
        print("请运行: npm install -g cloudconvert-cli")
        sys.exit(1)
    
    print(f"\n正在转换: {pptx_path}")
    print(f"输出目录: {output_dir}")
    print(f"图片格式: {image_format}")
    print("-" * 50)
    
    # 构建输出文件名模式 (page number)
    output_pattern = output_dir / f"slide_{pptx_path.stem}_{{page}}.{image_format}"
    
    # 使用 CloudConvert CLI 转换
    # cloudconvert convert -f jpg --apikey xxx input.pptx
    cmd = [
        'cloudconvert', 'convert',
        '-f', image_format,
        '--apikey', api_key,
        '--outputdir', str(output_dir),
        str(pptx_path)
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    print("转换中，请耐心等待...")
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
    except subprocess.CalledProcessError as e:
        print(f"转换失败: {e}")
        print(f"错误输出: {e.stderr}")
        sys.exit(1)
    
    # 获取生成的图片文件
    # CloudConvert 生成格式: {filename}-{page}.{format}
    image_files = list(output_dir.glob(f"{pptx_path.stem}-*.{image_format}"))
    image_files.sort(key=natural_sort_key)
    
    print(f"\n转换完成！共生成 {len(image_files)} 张图片")
    return image_files


def create_3x3_collage(image_paths, output_path, padding=10, bg_color=(255, 255, 255)):
    """
    创建3x3九宫格拼图
    
    Args:
        image_paths: 图片路径列表（最多9张）
        output_path: 输出路径
        padding: 图片之间的间距
        bg_color: 背景颜色
    """
    image_paths = image_paths[:9]
    
    if len(image_paths) == 0:
        print("没有图片需要处理")
        return
    
    images = []
    for path in image_paths:
        try:
            img = Image.open(path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            images.append(img)
        except Exception as e:
            print(f"无法打开图片 {path}: {e}")
    
    if len(images) == 0:
        print("没有有效的图片")
        return
    
    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)
    
    cell_width = max_width
    cell_height = max_height
    
    output_width = cell_width * 3 + padding * 4
    output_height = cell_height * 3 + padding * 4
    
    collage = Image.new('RGB', (output_width, output_height), bg_color)
    
    for idx, img in enumerate(images):
        row = idx // 3
        col = idx % 3
        
        x = padding + col * (cell_width + padding)
        y = padding + row * (cell_height + padding)
        
        img_resized = img.resize((cell_width, cell_height), Image.Resampling.LANCZOS)
        collage.paste(img_resized, (x, y))
    
    collage.save(output_path, quality=95)
    print(f"[OK] 已保存: {output_path} ({len(images)} 张图片)")
    
    return collage


def process_pptx_to_grid(pptx_path, output_base_dir=None):
    """
    处理 PPT 文件：先转图片，再生成九宫格
    
    Args:
        pptx_path: PPTX 文件路径
        output_base_dir: 基础输出目录
    """
    pptx_path = Path(pptx_path)
    
    if not pptx_path.exists():
        print(f"错误: 文件不存在: {pptx_path}")
        sys.exit(1)
    
    if pptx_path.suffix.lower() not in ['.pptx', '.ppt']:
        print(f"错误: 不支持的文件格式: {pptx_path.suffix}")
        sys.exit(1)
    
    # 设置输出目录
    if output_base_dir is None:
        output_base_dir = pptx_path.stem + "_output"
    output_base = Path(output_base_dir)
    
    # 1. 创建临时目录存放转换后的图片
    temp_images_dir = output_base / "slides"
    temp_images_dir.mkdir(parents=True, exist_ok=True)
    
    # 2. 使用 CloudConvert 转换 PPT 为图片
    image_files = convert_pptx_with_cloudconvert(
        pptx_path, 
        temp_images_dir, 
        image_format="jpg"
    )
    
    if len(image_files) == 0:
        print("错误: 没有生成任何图片")
        sys.exit(1)
    
    # 3. 创建九宫格拼图
    print("\n" + "=" * 50)
    print("开始生成九宫格拼图...")
    print("=" * 50)
    
    grid_output_dir = output_base / "grids"
    grid_output_dir.mkdir(parents=True, exist_ok=True)
    
    batch_size = 9
    num_batches = (len(image_files) + batch_size - 1) // batch_size
    
    for i in range(num_batches):
        start_idx = i * batch_size
        end_idx = min(start_idx + batch_size, len(image_files))
        batch_files = image_files[start_idx:end_idx]
        
        output_name = grid_output_dir / f"grid_{i+1:02d}.jpg"
        print(f"\n处理第 {i+1}/{num_batches} 组: 幻灯片 {start_idx+1}-{end_idx}")
        
        create_3x3_collage(batch_files, output_name)
    
    print("\n" + "=" * 50)
    print(f"全部完成！")
    print(f"  - 幻灯片图片: {temp_images_dir}")
    print(f"  - 九宫格拼图: {grid_output_dir}")
    print(f"  - 共生成 {num_batches} 个九宫格")
    print("=" * 50)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n参数:")
        print("  <pptx文件路径>  要转换的 PowerPoint 文件")
        print("  [输出目录]      可选，默认为文件名 + '_output'")
        sys.exit(1)
    
    pptx_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    process_pptx_to_grid(pptx_file, output_dir)


if __name__ == "__main__":
    main()
