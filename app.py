#!/usr/bin/env python3
"""
PPT to Grid - Streamlit Web App
将 PowerPoint 文件转换为九宫格拼图的网页应用
使用方法: uv run streamlit run app.py
"""

import os
import re
import zipfile
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st
from dotenv import load_dotenv
from PIL import Image

load_dotenv()


def natural_sort_key(s):
    """自然排序，正确处理数字"""
    s = str(s)
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r"([0-9]+)", s)]


def convert_pptx_with_cloudconvert(pptx_path: Path, output_dir: Path, api_key: str):
    """使用 CloudConvert API 将 PPT 转换为图片"""
    import cloudconvert
    import requests

    cloudconvert.configure(api_key=api_key)
    output_dir.mkdir(parents=True, exist_ok=True)

    job = cloudconvert.Job.create(
        payload={
            "tasks": {
                "import-file": {"operation": "import/upload"},
                "convert-file": {
                    "operation": "convert",
                    "input": "import-file",
                    "input_format": pptx_path.suffix.lstrip(".").lower(),
                    "output_format": "jpg",
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

    upload_task = job["tasks"][0]
    upload_url = upload_task["result"]["form"]["url"]
    upload_params = upload_task["result"]["form"]["parameters"]

    with open(pptx_path, "rb") as f:
        files = {"file": (pptx_path.name, f)}
        response = requests.post(upload_url, data=upload_params, files=files, timeout=300)
        response.raise_for_status()

    job = cloudconvert.Job.wait(id=job["id"])

    export_task = next((task for task in job["tasks"] if task["operation"] == "export/url"), None)
    if not export_task or not export_task.get("result"):
        raise Exception("转换失败")

    files_info = export_task["result"].get("files", [])
    if not files_info:
        raise Exception("没有生成任何文件")

    image_files = []
    for i, file_info in enumerate(files_info, 1):
        url = file_info["url"]
        output_path = output_dir / f"slide_{i:03d}.jpg"
        response = requests.get(url, timeout=300)
        response.raise_for_status()
        with open(output_path, "wb") as f:
            f.write(response.content)
        image_files.append(output_path)

    return sorted(image_files, key=natural_sort_key)


def create_3x3_collage(image_paths: list[Path], output_path: Path):
    """创建 3x3 九宫格拼图"""
    image_paths = image_paths[:9]
    if not image_paths:
        return None

    images = []
    for path in image_paths:
        img = Image.open(path)
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)

    if not images:
        return None

    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)
    padding, bg_color = 10, (255, 255, 255)
    cell_width, cell_height = max_width, max_height

    output_width = cell_width * 3 + padding * 4
    output_height = cell_height * 3 + padding * 4
    collage = Image.new("RGB", (output_width, output_height), bg_color)

    for idx, img in enumerate(images):
        row, col = idx // 3, idx % 3
        x = padding + col * (cell_width + padding)
        y = padding + row * (cell_height + padding)
        img_resized = img.resize((cell_width, cell_height), Image.Resampling.LANCZOS)
        collage.paste(img_resized, (x, y))

    collage.save(output_path, quality=95)
    return output_path


def create_zip_from_grids(grid_files: list[Path]) -> BytesIO:
    """将九宫格文件打包成 zip"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for grid_file in grid_files:
            zip_file.write(grid_file, grid_file.name)
    zip_buffer.seek(0)
    return zip_buffer


def process_pptx(pptx_path: Path, api_key: str, progress_bar, status_text):
    """处理 PPTX 文件的完整流程"""
    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        slides_dir = temp_path / "slides"
        grids_dir = temp_path / "grids"
        grids_dir.mkdir(parents=True, exist_ok=True)

        status_text.text("正在转换 PPT 为图片...")
        progress_bar.progress(10)

        image_files = convert_pptx_with_cloudconvert(pptx_path, slides_dir, api_key)
        if not image_files:
            raise Exception("没有生成任何图片")

        status_text.text(f"正在生成九宫格拼图 (共 {len(image_files)} 张幻灯片)...")
        progress_bar.progress(40)

        batch_size, num_batches = 9, (len(image_files) + 8) // 9
        grid_files = []
        for i in range(num_batches):
            start_idx = i * batch_size
            end_idx = min(start_idx + batch_size, len(image_files))
            batch_files = image_files[start_idx:end_idx]
            output_name = grids_dir / f"grid_{i+1:02d}.jpg"
            create_3x3_collage(batch_files, output_name)
            grid_files.append(output_name)
            progress_bar.progress(40 + int((i + 1) / num_batches * 40))

        status_text.text("正在打包文件...")
        progress_bar.progress(90)
        zip_buffer = create_zip_from_grids(grid_files)
        progress_bar.progress(100)
        status_text.text("完成！")

        return grid_files, zip_buffer


def main():
    st.set_page_config(page_title="PPT to Grid", page_icon="🖼️", layout="centered")
    st.title("🖼️ PPT to Grid")
    st.markdown("将 PowerPoint 幻灯片转换为 3×3 九宫格拼图")

    with st.sidebar:
        st.header("⚙️ 配置")
        env_api_key = os.environ.get("CLOUDCONVERT_API_KEY", "")
        api_key = st.text_input("CloudConvert API Key", value=env_api_key, type="password",
                                help="从 https://cloudconvert.com/dashboard/api 获取")
        if not api_key:
            st.warning("请输入 CloudConvert API Key")
        st.divider()
        st.markdown("[获取 API Key](https://cloudconvert.com/dashboard/api)")

    st.header("📤 上传文件")
    uploaded_file = st.file_uploader("选择 PPT 文件", type=["pptx", "ppt"],
                                      help="支持 .pptx 和 .ppt 格式")

    if uploaded_file is not None and api_key:
        st.divider()
        st.header("🔄 转换进度")

        progress_bar = st.progress(0)
        status_text = st.empty()

        if st.button("开始转换", type="primary", use_container_width=True):
            try:
                with TemporaryDirectory() as temp_dir:
                    temp_pptx = Path(temp_dir) / uploaded_file.name
                    with open(temp_pptx, "wb") as f:
                        f.write(uploaded_file.getvalue())

                    grid_files, zip_buffer = process_pptx(temp_pptx, api_key, progress_bar, status_text)

                    st.success(f"✅ 转换完成！共生成 {len(grid_files)} 个九宫格")

                    st.divider()
                    st.header("📥 下载结果")

                    # 生成下载文件名
                    file_stem = Path(uploaded_file.name).stem
                    zip_filename = f"{file_stem}_grids.zip"

                    st.download_button(
                        label="📦 下载所有九宫格 (ZIP)",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True,
                    )

                    with st.expander("预览九宫格"):
                        for grid_file in grid_files:
                            st.image(str(grid_file), caption=grid_file.name, use_container_width=True)

            except Exception as e:
                st.error(f"❌ 转换失败: {str(e)}")
                progress_bar.empty()

    elif uploaded_file is not None and not api_key:
        st.error("⚠️ 请先输入 CloudConvert API Key")

    st.divider()
    st.caption("使用 CloudConvert API 进行转换")


if __name__ == "__main__":
    main()
