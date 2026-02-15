#!/usr/bin/env python3
"""打包3MF文件预览工具v1.4为exe文件"""
import subprocess
import os
import shutil
import time
from datetime import datetime

print("=" * 60)
print("开始打包3MF文件预览工具v1.4")
print("=" * 60)

# 设置工作目录
work_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(work_dir)
print(f"工作目录: {work_dir}")

# 检查必要文件
required_files = ['qt_3mf_previewer.py', 'figurine64.ico']
for file in required_files:
    if not os.path.exists(file):
        print(f"错误: 缺少必要文件 {file}")
        exit(1)

print("✓ 所有必要文件检查通过")

# 清理旧的构建文件
print("\n清理旧的构建文件...")
for dir_name in ['build', 'dist', '__pycache__']:
    if os.path.exists(dir_name):
        print(f"删除目录: {dir_name}")
        shutil.rmtree(dir_name)

# 清理spec文件
spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
for spec_file in spec_files:
    print(f"删除spec文件: {spec_file}")
    os.remove(spec_file)

print("✓ 清理完成")

# 运行PyInstaller
print("\n开始PyInstaller打包...")
print("这可能需要几分钟时间，请耐心等待...")

pyinstaller_cmd = [
    'python', '-m', 'PyInstaller',
    '--onefile',
    '--windowed',
    '--icon', 'figurine64.ico',
    '--name', '3MF预览工具v1.4',
    '--add-data', 'figurine64.ico;.',
    'qt_3mf_previewer.py'
]

print(f"执行命令: {' '.join(pyinstaller_cmd)}")

process = subprocess.Popen(
    pyinstaller_cmd,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True,
    encoding='utf-8',
    errors='replace'
)

# 实时显示输出
while True:
    output = process.stdout.readline()
    if output == '' and process.poll() is not None:
        break
    if output:
        print(output.strip())

# 获取返回码
return_code = process.poll()

if return_code != 0:
    print(f"\n错误: 打包失败，返回码: {return_code}")
    stderr = process.stderr.read()
    if stderr:
        print("错误信息:")
        print(stderr)
    exit(1)

print("\n✓ PyInstaller打包完成")

# 等待文件写入完成
print("\n等待文件写入完成...")
time.sleep(5)

# 检查输出文件
print("\n检查输出文件...")
if os.path.exists('dist'):
    files = os.listdir('dist')
    print(f"dist目录内容: {files}")
    
    if files:
        for file in files:
            file_path = os.path.join('dist', file)
            if os.path.isfile(file_path):
                file_size = os.path.getsize(file_path)
                file_size_mb = file_size / (1024 * 1024)
                print(f"✓ 文件: {file}")
                print(f"  大小: {file_size_mb:.2f} MB ({file_size:,} 字节)")
                
                # 复制到当前目录
                dest_path = os.path.join(work_dir, file)
                shutil.copy2(file_path, dest_path)
                print(f"  已复制到: {dest_path}")
    else:
        print("错误: dist目录为空")
        exit(1)
else:
    print("错误: dist目录不存在")
    exit(1)

# 创建发布包
print("\n创建发布包...")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
release_dir = f"3MF预览工具v1.4_{timestamp}"
os.makedirs(release_dir, exist_ok=True)

# 复制exe文件到发布目录
for file in files:
    src = os.path.join('dist', file)
    dst = os.path.join(release_dir, file)
    shutil.copy2(src, dst)
    print(f"✓ 复制 {file} 到发布目录")

# 复制README
if os.path.exists('README.md'):
    shutil.copy2('README.md', os.path.join(release_dir, 'README.md'))
    print("✓ 复制 README.md 到发布目录")

# 创建zip文件
zip_filename = f"{release_dir}.zip"
print(f"\n创建zip文件: {zip_filename}")
shutil.make_archive(release_dir, 'zip', '.', release_dir)

# 获取zip文件大小
zip_size = os.path.getsize(zip_filename)
zip_size_mb = zip_size / (1024 * 1024)
print(f"✓ ZIP文件大小: {zip_size_mb:.2f} MB")

print("\n" + "=" * 60)
print("打包完成！")
print("=" * 60)
print(f"\n输出文件:")
print(f"1. EXE文件: {os.path.join(work_dir, files[0])}")
print(f"2. 发布目录: {os.path.join(work_dir, release_dir)}")
print(f"3. ZIP文件: {os.path.join(work_dir, zip_filename)}")
print(f"\n文件大小: {zip_size_mb:.2f} MB")
print("\n✓ 所有任务完成！")