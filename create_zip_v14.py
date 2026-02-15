#!/usr/bin/env python3
"""创建v1.4版本的压缩包"""
import os
import shutil
from datetime import datetime

print("=" * 60)
print("创建3MF文件预览工具v1.4压缩包")
print("=" * 60)

# 设置工作目录
work_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(work_dir)
print(f"工作目录: {work_dir}")

# 检查exe文件
exe_file = os.path.join('dist', '3MF预览工具v1.4.exe')
if not os.path.exists(exe_file):
    print(f"错误: 找不到exe文件 {exe_file}")
    exit(1)

file_size = os.path.getsize(exe_file)
file_size_mb = file_size / (1024 * 1024)
print(f"✓ 找到exe文件: {exe_file}")
print(f"  大小: {file_size_mb:.2f} MB ({file_size:,} 字节)")

# 创建发布包目录
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
release_dir = f"3MF预览工具v1.4_{timestamp}"
os.makedirs(release_dir, exist_ok=True)
print(f"\n创建发布目录: {release_dir}")

# 复制exe文件到发布目录
dest_exe = os.path.join(release_dir, '3MF预览工具v1.4.exe')
shutil.copy2(exe_file, dest_exe)
print(f"✓ 复制exe文件到发布目录")

# 复制README
if os.path.exists('README.md'):
    shutil.copy2('README.md', os.path.join(release_dir, 'README.md'))
    print("✓ 复制 README.md 到发布目录")

# 复制备份文件
if os.path.exists('3mf导出工具1.3.1备份.py'):
    shutil.copy2('3mf导出工具1.3.1备份.py', os.path.join(release_dir, '3mf导出工具1.3.1备份.py'))
    print("✓ 复制备份文件到发布目录")

# 创建zip文件
zip_filename = f"{release_dir}.zip"
print(f"\n创建zip文件: {zip_filename}")
shutil.make_archive(release_dir, 'zip', '.', release_dir)

# 获取zip文件大小
zip_size = os.path.getsize(zip_filename)
zip_size_mb = zip_size / (1024 * 1024)
print(f"✓ ZIP文件大小: {zip_size_mb:.2f} MB")

# 复制exe到当前目录
current_dir_exe = os.path.join(work_dir, '3MF预览工具v1.4.exe')
shutil.copy2(exe_file, current_dir_exe)
print(f"✓ 复制exe文件到当前目录")

print("\n" + "=" * 60)
print("压缩包创建完成！")
print("=" * 60)
print(f"\n输出文件:")
print(f"1. EXE文件: {current_dir_exe}")
print(f"2. 发布目录: {os.path.join(work_dir, release_dir)}")
print(f"3. ZIP文件: {os.path.join(work_dir, zip_filename)}")
print(f"\nEXE文件大小: {file_size_mb:.2f} MB")
print(f"ZIP文件大小: {zip_size_mb:.2f} MB")
print("\n✓ 所有任务完成！")