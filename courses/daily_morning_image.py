#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
早安问候图片自动生成脚本
每天7:00运行，生成带有"早安你好！"文字的图片
"""
import os
import sys
sys.path.insert(0, '/opt/homebrew/lib/node_modules/openclaw/node_modules')

from PIL import Image, ImageDraw, ImageFont
import subprocess
import json
from datetime import datetime

# Today's date for variety
TODAY = datetime.now().strftime("%Y年%m月%d日")

# Different prompts for variety
PROMPTS = [
    "Beautiful sunrise over Chinese mountains, peaceful morning scenery, traditional landscape art style, soft golden light, misty valleys, elegant and serene atmosphere, high quality",
    "Cherry blossoms in morning light, Japanese garden style, delicate flowers, peaceful pond, traditional Asian aesthetics, soft pink and white colors, beautiful morning scene",
    "Lotus pond at sunrise, traditional Chinese painting style, pink lotus flowers, morning dew, peaceful water surface, elegant natural beauty",
    "Ancient Chinese pavilion at sunrise, mountain backdrop, traditional architecture, soft morning colors, peaceful and scenic, beautiful landscape",
    "Bamboo forest morning, traditional Chinese landscape, green bamboo, soft sunlight filtering through, serene and peaceful atmosphere",
    "Mountain stream at dawn, traditional Chinese landscape painting style, flowing water, rocks, morning mist, peaceful natural scenery",
    "Plum blossoms in early morning, traditional Chinese art style, white and pink flowers, peaceful winter scene, elegant beauty",
    "Willow trees by lake at sunrise, traditional Chinese landscape, drooping willows, calm water, soft morning light, peaceful scenery",
]

def generate_image():
    """Generate morning greeting image with AI"""
    # Use date to select prompt for variety
    import random
    prompt = random.choice(PROMPTS)
    
    # Call image generation
    result = subprocess.run([
        'curl', '-s', '-X', 'POST',
        'https://api.minimax.chat/v1/image_generation',
        '-H', 'Authorization: Bearer ' + os.environ.get('MINIMAX_API_KEY', ''),
        '-H', 'Content-Type: application/json',
        '-d', json.dumps({
            "model": "image-01",
            "prompt": prompt,
            "aspect_ratio": "1:1"
        })
    ], capture_output=True, text=True, timeout=30)
    
    return result.stdout

def create_morning_image():
    """Create morning greeting image with text overlay"""
    output_dir = "/Users/mac/.openclaw/media/outbound"
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate or fetch base image
    # For now, we'll create a beautiful gradient image as base
    img = Image.new('RGB', (1024, 1024), color=(135, 206, 235))  # Sky blue
    
    # Create a nice gradient sky
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    for y in range(512):
        color_val = int(135 + (200 - 135) * (y / 512))
        draw.rectangle([(0, y), (1024, y+1)], fill=(color_val, 180 + y//4, 100 + y//8))
    
    # Add a simple sun
    for i in range(100):
        x, y = 750 - i//2, 300 - i//2
        r = 80 + i
        if r > 0:
            draw.ellipse([(x-r, y-r), (x+r, y+r)], fill=(255, 200 - i, 100 - i//2))
    
    # Draw mountains silhouette
    draw.polygon([(0, 600), (200, 450), (400, 550), (600, 400), (800, 520), (1024, 480), (1024, 1024), (0, 1024)], fill=(60, 90, 60))
    draw.polygon([(0, 700), (300, 550), (500, 650), (700, 500), (900, 620), (1024, 580), (1024, 1024), (0, 1024)], fill=(40, 70, 40))
    
    # Font settings - try Chinese font
    font_size = 80
    font_paths = [
        "/System/Library/Fonts/STKaiti.ttc",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
    ]
    
    font = None
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                font = ImageFont.truetype(fp, font_size)
                print(f"Using font: {fp}")
                break
            except:
                continue
    
    if font is None:
        font = ImageFont.load_default()
    
    # Draw text with shadow
    text = "早安你好！"
    
    # Get text bounding box
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    x = (1024 - text_width) // 2
    y = 850
    
    # Shadow
    draw.text((x+4, y+4), text, font=font, fill=(0, 0, 0))
    # Main text
    draw.text((x, y), text, font=font, fill=(255, 255, 255))
    
    # Save
    output_path = os.path.join(output_dir, f"早安问候_{datetime.now().strftime('%Y%m%d')}.jpg")
    img.save(output_path, quality=95)
    print(f"Saved: {output_path}")
    return output_path

if __name__ == "__main__":
    try:
        path = create_morning_image()
        print(f"SUCCESS: {path}")
    except Exception as e:
        print(f"ERROR: {e}")
