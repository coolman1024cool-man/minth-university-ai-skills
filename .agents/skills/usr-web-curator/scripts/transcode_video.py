# -*- coding: utf-8 -*-
"""
USR Web Curator - Video Transcoder
自動將原始大型 MP4 影音壓縮轉碼為 720p H.264/AAC 串流輕量規格 (CRF 28, 96k audio)
"""
import os
import sys
import subprocess

def transcode(input_file, output_file):
    if not os.path.exists(input_file):
        print(f"錯誤：找不到輸入檔案 {input_file}")
        return False
    
    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
    
    cmd = [
        "ffmpeg", "-y",
        "-i", input_file,
        "-vf", "scale=-2:720",
        "-c:v", "libx264",
        "-crf", "28",
        "-preset", "fast",
        "-c:a", "aac",
        "-b:a", "96k",
        output_file
    ]
    print(f"開始轉碼: {os.path.basename(input_file)} -> {output_file}")
    res = subprocess.run(cmd, capture_output=True)
    if res.returncode == 0:
        orig_mb = os.path.getsize(input_file) / (1024 * 1024)
        out_mb = os.path.getsize(output_file) / (1024 * 1024)
        print(f"✅ 轉碼成功！體積由 {orig_mb:.1f} MB 瘦身至 {out_mb:.1f} MB")
        return True
    else:
        print(f"❌ 轉碼失敗: {res.stderr.decode('utf-8', errors='ignore')}")
        return False

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python transcode_video.py <輸入影片路徑> <輸出影片路徑>")
    else:
        transcode(sys.argv[1], sys.argv[2])
