from PIL import Image
import os
import re
from pathlib import Path

def natural_sort_key(s):
    """自然排序，正确处理数字"""
    return [int(text) if text.isdigit() else text.lower() 
            for text in re.split('([0-9]+)', s)]

def create_3x3_collage(image_paths, output_path, padding=10, bg_color=(255, 255, 255)):
    """
    创建3x3九宫格拼图
    
    Args:
        image_paths: 图片路径列表（最多9张）
        output_path: 输出路径
        padding: 图片之间的间距
        bg_color: 背景颜色
    """
    # 只取前9张
    image_paths = image_paths[:9]
    
    if len(image_paths) == 0:
        print("没有图片需要处理")
        return
    
    # 打开所有图片
    images = []
    for path in image_paths:
        try:
            img = Image.open(path)
            # 转换为RGB模式（处理RGBA图片）
            if img.mode != 'RGB':
                img = img.convert('RGB')
            images.append(img)
        except Exception as e:
            print(f"无法打开图片 {path}: {e}")
    
    if len(images) == 0:
        print("没有有效的图片")
        return
    
    # 计算目标单元格大小（使用所有图片中的最大宽高，或者可以统一缩放）
    # 这里我们找到合适的尺寸使所有图片能整齐排列
    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)
    
    # 统一使用相同的单元格大小
    cell_width = max_width
    cell_height = max_height
    
    # 计算输出图片的尺寸
    output_width = cell_width * 3 + padding * 4
    output_height = cell_height * 3 + padding * 4
    
    # 创建背景画布
    collage = Image.new('RGB', (output_width, output_height), bg_color)
    
    # 填充图片
    for idx, img in enumerate(images):
        row = idx // 3  # 行号 0, 1, 2
        col = idx % 3   # 列号 0, 1, 2
        
        # 计算位置
        x = padding + col * (cell_width + padding)
        y = padding + row * (cell_height + padding)
        
        # 等比例缩放图片以适应单元格
        img_resized = img.resize((cell_width, cell_height), Image.Resampling.LANCZOS)
        
        # 粘贴到画布
        collage.paste(img_resized, (x, y))
    
    # 保存结果
    collage.save(output_path, quality=95)
    print(f"已保存: {output_path} ({len(images)} 张图片)")
    
    return collage

def main():
    # 获取当前目录
    current_dir = Path(".")
    
    # 获取所有jpg图片并按文件名排序
    image_files = list(current_dir.glob("*.jpg"))
    image_files = [f for f in image_files if f.is_file()]
    image_files.sort(key=lambda x: natural_sort_key(x.name))
    
    print(f"找到 {len(image_files)} 张图片")
    
    if len(image_files) == 0:
        print("当前目录没有jpg图片")
        return
    
    # 按每9张分组处理
    batch_size = 9
    num_batches = (len(image_files) + batch_size - 1) // batch_size
    
    for i in range(num_batches):
        start_idx = i * batch_size
        end_idx = min(start_idx + batch_size, len(image_files))
        batch_files = image_files[start_idx:end_idx]
        
        output_name = f"collage_3x3_{i+1:02d}.jpg"
        print(f"\n处理第 {i+1}/{num_batches} 组: {start_idx+1}-{end_idx} 张图片")
        
        create_3x3_collage(batch_files, output_name)
    
    print(f"\n完成！共生成 {num_batches} 个九宫格图片")

if __name__ == "__main__":
    main()
