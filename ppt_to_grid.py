#!/usr/bin/env python3
"""
PPT 转九宫格拼图工具

工作流程:
1. 使用 CloudConvert 将 PPT 转换为图片
2. 将图片按顺序拼成 3x3 九宫格

配置:
    创建 .env 文件并设置 CLOUDCONVERT_API_KEY:
    CLOUDCONVERT_API_KEY=your_api_key

    或设置环境变量:
    export CLOUDCONVERT_API_KEY=your_api_key

使用方法:
    uv run ppt_to_grid.py <pptx文件路径> [输出目录]

示例:
    uv run ppt_to_grid.py "presentation.pptx"
    uv run ppt_to_grid.py "presentation.pptx" ./output
"""

import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from PIL import Image

# 加载 .env 文件
load_dotenv()


def natural_sort_key(s: str | Path) -> list:
    """自然排序，正确处理数字"""
    s = str(s)
    return [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r"([0-9]+)", s)
    ]


def convert_pptx_with_cloudconvert(
    pptx_path: Path, output_dir: Path, image_format: str = "jpg"
) -> list[Path]:
    """
    使用 CloudConvert API 将 PPT 转换为图片

    Args:
        pptx_path: PPTX 文件路径
        output_dir: 输出目录
        image_format: 图片格式 (jpg 或 png)

    Returns:
        生成的图片路径列表
    """
    import cloudconvert

    api_key = os.environ.get("CLOUDCONVERT_API_KEY")
    if not api_key:
        print("错误: 未设置 CLOUDCONVERT_API_KEY")
        print("\n请通过以下方式之一设置:")
        print("1. 创建 .env 文件:")
        print("   CLOUDCONVERT_API_KEY=your_api_key")
        print("2. 设置环境变量:")
        print("   export CLOUDCONVERT_API_KEY=your_api_key")
        print("\n获取 API Key: https://cloudconvert.com/dashboard/api")
        sys.exit(1)

    cloudconvert.configure(api_key=api_key)

    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n正在转换: {pptx_path}")
    print(f"输出目录: {output_dir}")
    print(f"图片格式: {image_format}")
    print("-" * 50)

    try:
        # 创建转换任务: PPTX -> JPG/PNG
        job = cloudconvert.Job.create(
            payload={
                "tasks": {
                    "import-file": {
                        "operation": "import/upload",
                    },
                    "convert-file": {
                        "operation": "convert",
                        "input": "import-file",
                        "input_format": pptx_path.suffix.lstrip(".").lower(),
                        "output_format": image_format,
                        "engine": "libreoffice",
                    },
                    "export-file": {
                        "operation": "export/url",
                        "input": "convert-file",
                        "inline": False,
                        "archive_multiple_files": False,
                    },
                }
            }
        )

        # 上传文件
        upload_task = job["tasks"][0]
        upload_url = upload_task["result"]["form"]["url"]
        upload_params = upload_task["result"]["form"]["parameters"]

        with open(pptx_path, "rb") as f:
            import requests

            files = {"file": (pptx_path.name, f)}
            response = requests.post(
                upload_url, data=upload_params, files=files, timeout=300
            )
            response.raise_for_status()

        print("文件已上传，等待转换...")

        # 等待任务完成
        job = cloudconvert.Job.wait(id=job["id"])

        # 获取下载链接
        export_task = next(
            (task for task in job["tasks"] if task["operation"] == "export/url"),
            None,
        )

        if not export_task or not export_task.get("result"):
            print("错误: 转换失败")
            sys.exit(1)

        files = export_task["result"].get("files", [])
        if not files:
            print("错误: 没有生成任何文件")
            sys.exit(1)

        # 下载文件
        image_files = []
        for i, file_info in enumerate(files, 1):
            url = file_info["url"]
            filename = file_info["filename"]
            output_path = output_dir / f"slide_{i:03d}.{image_format}"

            response = requests.get(url, timeout=300)
            response.raise_for_status()

            with open(output_path, "wb") as f:
                f.write(response.content)

            image_files.append(output_path)
            print(f"[{i}/{len(files)}] 已下载: {output_path.name}")

        print(f"\n转换完成！共生成 {len(image_files)} 张图片")
        return sorted(image_files, key=natural_sort_key)

    except Exception as e:
        print(f"转换失败: {e}")
        sys.exit(1)


def create_3x3_collage(
    image_paths: list[Path],
    output_path: Path,
    padding: int = 10,
    bg_color: tuple[int, int, int] = (255, 255, 255),
) -> Image.Image | None:
    """
    创建 3x3 九宫格拼图

    Args:
        image_paths: 图片路径列表（最多 9 张）
        output_path: 输出路径
        padding: 图片之间的间距
        bg_color: 背景颜色

    Returns:
        拼接后的图片对象
    """
    image_paths = image_paths[:9]

    if not image_paths:
        print("没有图片需要处理")
        return None

    images = []
    for path in image_paths:
        try:
            img = Image.open(path)
            if img.mode != "RGB":
                img = img.convert("RGB")
            images.append(img)
        except Exception as e:
            print(f"无法打开图片 {path}: {e}")

    if not images:
        print("没有有效的图片")
        return None

    # 统一单元格大小
    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)

    cell_width = max_width
    cell_height = max_height

    # 计算画布大小
    output_width = cell_width * 3 + padding * 4
    output_height = cell_height * 3 + padding * 4

    # 创建画布
    collage = Image.new("RGB", (output_width, output_height), bg_color)

    # 填充图片
    for idx, img in enumerate(images):
        row = idx // 3
        col = idx % 3

        x = padding + col * (cell_width + padding)
        y = padding + row * (cell_height + padding)

        img_resized = img.resize((cell_width, cell_height), Image.Resampling.LANCZOS)
        collage.paste(img_resized, (x, y))

    collage.save(output_path, quality=95)
    print(f"  [OK] 已保存: {output_path.name} ({len(images)} 张图片)")

    return collage


def process_pptx_to_grid(pptx_path: Path, output_base_dir: Path | None = None) -> None:
    """
    处理 PPT 文件：先转图片，再生成九宫格

    Args:
        pptx_path: PPTX 文件路径
        output_base_dir: 基础输出目录
    """
    if not pptx_path.exists():
        print(f"错误: 文件不存在: {pptx_path}")
        sys.exit(1)

    if pptx_path.suffix.lower() not in [".pptx", ".ppt"]:
        print(f"错误: 不支持的文件格式: {pptx_path.suffix}")
        print("只支持 .pptx 和 .ppt 格式")
        sys.exit(1)

    # 设置输出目录
    if output_base_dir is None:
        output_base_dir = Path(pptx_path.stem + "_output")
    else:
        output_base_dir = Path(output_base_dir)

    # 创建临时目录存放转换后的图片
    temp_images_dir = output_base_dir / "slides"

    # 使用 CloudConvert 转换 PPT 为图片
    image_files = convert_pptx_with_cloudconvert(
        pptx_path, temp_images_dir, image_format="jpg"
    )

    if not image_files:
        print("错误: 没有生成任何图片")
        sys.exit(1)

    # 创建九宫格拼图
    print("\n" + "=" * 50)
    print("开始生成九宫格拼图...")
    print("=" * 50)

    grid_output_dir = output_base_dir / "grids"
    grid_output_dir.mkdir(parents=True, exist_ok=True)

    batch_size = 9
    num_batches = (len(image_files) + batch_size - 1) // batch_size

    for i in range(num_batches):
        start_idx = i * batch_size
        end_idx = min(start_idx + batch_size, len(image_files))
        batch_files = image_files[start_idx:end_idx]

        output_name = grid_output_dir / f"grid_{i+1:02d}.jpg"
        print(f"\n第 {i+1}/{num_batches} 组: 幻灯片 {start_idx+1}-{end_idx}")

        create_3x3_collage(batch_files, output_name)

    print("\n" + "=" * 50)
    print("全部完成！")
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

    pptx_file = Path(sys.argv[1])
    output_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else None

    process_pptx_to_grid(pptx_file, output_dir)


if __name__ == "__main__":
    main()
