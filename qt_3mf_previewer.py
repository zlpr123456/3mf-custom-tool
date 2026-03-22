#!/usr/bin/env python3
"""
3MF文件预览工具
版本: v1.5
功能:
1. 打开和预览3MF文件
2. 查看文件中的G-code内容
3. 支持多打印盘3MF文件
4. 显示文件的元数据信息
5. 添加自定义G-code代码
6. 导出修改后的3MF文件
7. 导出最终的G-code文件
"""

import os
import sys
import json
import re
import tempfile
import zipfile
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QTextEdit, QFileDialog,
    QVBoxLayout, QWidget, QComboBox, QLabel, QGroupBox, QHBoxLayout,
    QSplitter, QMessageBox, QInputDialog, QProgressDialog, QDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QIcon

class ThreeMfPreviewer(QMainWindow):
    """3MF文件预览器"""
    
    def __init__(self):
        """初始化预览器"""
        super().__init__()
        
        # 初始化变量
        self.file_info = ""
        self.metadata = {}
        self.three_mf_files = []
        self.preview_image_path = ""
        self.gcode_content = ""
        self.gcode_files = []
        self.current_file_index = -1
        self.file_data = []  # 存储每个文件的数据：(file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode)
        self.custom_gcode = ""
        # G-code缓存，用于快速打开同一个文件的G-code
        self.gcode_cache = {}
        
        # 初始化用户界面
        self.init_ui()
    
    def init_ui(self):
        """初始化用户界面"""
        # 初始化用户界面
        self.setWindowTitle("A1自动换盘")
        self.setGeometry(100, 100, 1000, 700)
        
        # 设置窗口图标
        import os
        icon_path = os.path.join(os.path.dirname(__file__), "figurine64.ico")
        if os.path.exists(icon_path):
            from PyQt5.QtGui import QIcon
            self.setWindowIcon(QIcon(icon_path))
        
        # 设置现代风格
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                margin-top: 10px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #333333;
                font-weight: bold;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QComboBox {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 5px;
                min-width: 200px;
                font-size: 14px;
            }
            QTextEdit {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                font-family: "SimSun", "Courier New";
                font-size: 14px;
            }
            QLabel {
                font-size: 14px;
            }
        """)
        
        # 创建主部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建按钮区，分成两排
        button_container = QWidget()
        button_layout = QVBoxLayout(button_container)
        
        # 第一排按钮
        top_button_layout = QHBoxLayout()
        
        # 创建打开文件按钮
        self.open_button = QPushButton("打开3MF文件...")
        self.open_button.clicked.connect(self.open_file)
        top_button_layout.addWidget(self.open_button)
        
        # 创建打开多个文件按钮
        self.open_multiple_button = QPushButton("打开多个3MF文件...")
        self.open_multiple_button.clicked.connect(self.open_multiple_files)
        top_button_layout.addWidget(self.open_multiple_button)
        
        # 创建导出文件按钮
        self.export_button = QPushButton("导出3MF文件...")
        self.export_button.clicked.connect(self.export_file)
        top_button_layout.addWidget(self.export_button)
        
        # 创建导出最终文件按钮，并设置美观的蓝色底色和悬停效果
        self.export_final_button = QPushButton("导出最终文件...")
        self.export_final_button.clicked.connect(self.export_final_file)
        self.export_final_button.setStyleSheet("""
            QPushButton {
                background-color: #1976D2;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
        """)
        top_button_layout.addWidget(self.export_final_button)
        
        # 第二排按钮
        bottom_button_layout = QHBoxLayout()
        
        # 创建添加G-code按钮
        self.add_gcode_button = QPushButton("添加自定义G-code...")
        self.add_gcode_button.clicked.connect(self.add_gcode)
        bottom_button_layout.addWidget(self.add_gcode_button)
        
        # 创建参数搜索按钮
        self.search_button = QPushButton("参数搜索...")
        self.search_button.clicked.connect(self.search_parameters)
        bottom_button_layout.addWidget(self.search_button)
        
        # 创建换色代码对比按钮
        self.color_change_compare_button = QPushButton("换色代码对比")
        self.color_change_compare_button.clicked.connect(self.compare_color_change_codes)
        bottom_button_layout.addWidget(self.color_change_compare_button)
        
        # 创建显示G-code按钮
        self.show_gcode_button = QPushButton("显示G-code")
        self.show_gcode_button.clicked.connect(self.show_gcode)
        bottom_button_layout.addWidget(self.show_gcode_button)
        
        # 创建清空按钮，并设置美观的红色底色和悬停效果
        self.clear_button = QPushButton("清空文件")
        self.clear_button.clicked.connect(self.clear_files)
        self.clear_button.setStyleSheet("""
            QPushButton {
                background-color: #E53935;
                color: white;
                border: none;
                padding: 8px 16px;
                text-align: center;
                text-decoration: none;
                font-size: 14px;
                margin: 4px 2px;
                border-radius: 4px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #C62828;
            }
            QPushButton:pressed {
                background-color: #B71C1C;
            }
        """)
        bottom_button_layout.addWidget(self.clear_button)
        
        # 创建关于按钮
        self.about_button = QPushButton("关于")
        self.about_button.clicked.connect(self.about)
        bottom_button_layout.addWidget(self.about_button)
        
        # 创建退出按钮
        self.exit_button = QPushButton("退出")
        self.exit_button.clicked.connect(self.close)
        bottom_button_layout.addWidget(self.exit_button)
        
        # 将两排按钮添加到容器布局，并设置间距
        button_layout.addLayout(top_button_layout)
        # 设置两排按钮之间的间距，使用较小的值
        button_layout.setSpacing(5)
        button_layout.addLayout(bottom_button_layout)
        
        # 添加按钮容器到主布局，并设置其固定高度
        button_container.setFixedHeight(100)  # 设置固定高度
        main_layout.addWidget(button_container)
        
        # 创建文件选择下拉框
        self.file_var = ""
        self.file_selector = QComboBox()
        self.file_selector.currentIndexChanged.connect(self.on_file_selected)
        main_layout.addWidget(self.file_selector)
        
        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        
        # 创建左侧预览区
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # 创建预览图组
        preview_group = QGroupBox("预览图")
        preview_layout = QVBoxLayout(preview_group)
        
        # 创建预览图标签
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(300, 300)
        self.preview_label.setStyleSheet("background-color: #f9f9f9; border: 1px solid #d0d0d0;")
        preview_layout.addWidget(self.preview_label)
        
        left_layout.addWidget(preview_group)
        
        # 添加左侧部件到分割器
        splitter.addWidget(left_widget)
        
        # 创建右侧参数区
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 创建文件信息组
        info_group = QGroupBox("文件信息")
        info_layout = QVBoxLayout(info_group)
        
        # 创建参数显示文本框
        self.info_text = QTextEdit()
        self.info_text.setReadOnly(True)
        self.info_text.setPlainText("3MF文件预览工具\n\n点击'打开3MF文件...'按钮选择要预览的3MF文件。")
        info_layout.addWidget(self.info_text)
        
        right_layout.addWidget(info_group)
        
        # 添加右侧部件到分割器
        splitter.addWidget(right_widget)
        
        # 设置分割器比例
        splitter.setSizes([350, 650])
        
        # 添加分割器到主布局，并设置其为可扩展的
        main_layout.addWidget(splitter)
        # 设置分割器为可伸缩的，使其在窗口放大时能够动态缩放
        main_layout.setStretch(2, 1)  # 分割器占据剩余空间
    
    def open_file(self):
        """打开单个3MF文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "打开3MF文件", "", "3MF Files (*.3mf);;All Files (*.*)"
        )
        
        if file_path:
            print(f"打开文件: {file_path}")
            
            # 显示进度对话框
            progress = QProgressDialog("正在解析文件，请稍候...", "取消", 0, 0, self)
            progress.setWindowTitle("解析文件")
            progress.setWindowModality(Qt.WindowModal)
            progress.show()
            
            # 直接调用解析方法
            try:
                plate_data = self.parse_three_mf_file(file_path, 0)
                progress.close()
                file_name = os.path.basename(file_path)
                
                if plate_data:
                    # 更新当前文件信息
                    self.current_file = plate_data[0][0]
                    # 清空之前的文件列表和数据
                    self.three_mf_files = [self.current_file]
                    self.file_data = plate_data
                    self.current_file_index = 0
                    # 清空并重新配置文件选择下拉框
                    self.file_selector.clear()
                    base_file_name = os.path.basename(self.current_file)
                    # 为每个打印盘添加一个选项
                    for i, plate in enumerate(plate_data):
                        if len(plate_data) > 1:
                            self.file_selector.addItem(f"{base_file_name} - Plate {i+1}")
                        else:
                            self.file_selector.addItem(base_file_name)
                    # 设置默认选项为第一个打印盘
                    self.file_selector.setCurrentIndex(0)
                    # 显示第一个打印盘的文件信息和预览图
                    if plate_data:
                        file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode = plate_data[0]
                        self.file_info = file_info
                        self.preview_image_path = preview_image_path
                        self.gcode_content = gcode_content
                        self.metadata = metadata.copy()
                        self.custom_gcode = custom_gcode
                        self.display_file_info()
                        self.display_preview_image()
                        
                        # 检查是否有G-code内容
                        if not self.gcode_content:
                            QMessageBox.critical(self, "导入失败", f"【{file_name}】没有读取到gcode，请导入切片后的文件！")
                        else:
                            # 导入成功提示
                            QMessageBox.information(self, "导入成功", f"【{file_name}】导入成功。")
                else:
                    self.display_error()
                    QMessageBox.critical(self, "导入失败", f"【{file_name}】过大，无法解析。请调整文件内容。")
            except Exception as e:
                progress.close()
                file_name = os.path.basename(file_path)
                error_msg = "3MF文件信息\n"
                error_msg += "=" * 50 + "\n\n"
                error_msg += f"解析3MF文件时发生异常: {str(e)}\n"
                self.file_info = error_msg
                self.display_file_info()
                QMessageBox.critical(self, "导入失败", f"【{file_name}】过大，无法解析。请调整文件内容。")
    
    def open_multiple_files(self):
        """打开多个3MF文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "打开多个3MF文件", "", "3MF Files (*.3mf);;All Files (*.*)"
        )
        
        if file_paths:
            # 显示进度对话框
            progress = QProgressDialog("正在解析多个文件，请稍候...", "取消", 0, len(file_paths), self)
            progress.setWindowTitle("解析文件")
            progress.setWindowModality(Qt.WindowModal)
            progress.show()
            
            self.three_mf_files = file_paths
            self.file_data = []
            
            # 解析每个文件
            for i, file_path in enumerate(file_paths):
                progress.setValue(i)
                progress.setLabelText(f"正在解析文件 {i+1}/{len(file_paths)}...")
                
                # 同步解析每个文件（使用改进后的parse_three_mf_file方法）
                try:
                    plate_data = self.parse_three_mf_file(file_path, i)
                    if plate_data:
                        self.file_data.extend(plate_data)
                except Exception:
                    pass
                
                # 处理进度对话框的取消按钮
                if progress.wasCanceled():
                    break
            
            progress.setValue(len(file_paths))
            progress.close()
            
            # 清空文件选择下拉框
            self.file_selector.clear()
            
            # 为每个打印盘添加一个选项
            for i, plate in enumerate(self.file_data):
                try:
                    file_path = plate[0]
                    file_name = os.path.basename(file_path)
                    if len(self.file_data) > 1:
                        self.file_selector.addItem(f"{file_name} - Plate {i+1}")
                    else:
                        self.file_selector.addItem(file_name)
                except Exception:
                    pass
            
            # 设置默认选项为第一个打印盘
            if self.file_data:
                try:
                    self.file_selector.setCurrentIndex(0)
                    self.switch_file(0)
                except Exception:
                    pass
            else:
                QMessageBox.critical(self, "错误", "没有找到有效的打印盘数据！")
    
    def on_file_selected(self, index):
        """处理下拉框选择事件"""
        if index >= 0 and index < len(self.file_data):
            self.switch_file(index)
    
    def switch_file(self, index):
        """切换显示不同文件的内容"""
        if 0 <= index < len(self.file_data):
            self.current_file_index = index
            file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode = self.file_data[index]
            
            # 添加调试信息
            print(f"\n切换到文件: {os.path.basename(file_path)}")
            print(f"预览图路径: {preview_image_path}")
            print(f"预览图文件存在: {os.path.exists(preview_image_path) if preview_image_path else False}")
            print(f"元数据信息数量: {len(metadata)}")
            print(f"元数据信息: {metadata}")
            print(f"文件在file_data中的索引: {index}")
            
            # 更新当前文件信息
            self.current_file = file_path
            self.file_info = file_info
            self.preview_image_path = preview_image_path
            self.gcode_content = gcode_content
            self.metadata = metadata.copy()  # 确保使用元数据的副本
            self.custom_gcode = custom_gcode
            
            # 添加调试信息
            print(f"更新后的当前文件: {os.path.basename(self.current_file)}")
            print(f"更新后的预览图路径: {self.preview_image_path}")
            print(f"更新后的元数据数量: {len(self.metadata)}")
            
            # 强制显示文件信息和预览图
            print("开始显示文件信息...")
            self.display_file_info()
            print("开始显示预览图...")
            self.display_preview_image()
            print("预览图显示完成")
    
    def parse_three_mf_file(self, file_path, file_index):
        """解析3MF文件，支持多打印盘"""
        try:
            # 初始化变量
            file_size_bytes = 0
            file_size_mb = 0
            mod_time_str = ""
            
            # 存储每个打印盘的数据
            plate_data = []
            
            # 检查文件是否存在
            if not os.path.exists(file_path):
                return []
            
            # 检查文件是否为ZIP格式
            if not zipfile.is_zipfile(file_path):
                return []
            
            try:
                # 打开ZIP文件
                with zipfile.ZipFile(file_path, 'r') as zf:
                    try:
                        # 获取所有文件列表
                        file_list = zf.namelist()
                        
                        # 收集所有G-code文件和预览图
                        gcode_files = []
                        preview_files = []
                        metadata_files = []
                        
                        for file_name in file_list:
                            try:
                                if file_name.endswith('.gcode'):
                                    gcode_files.append(file_name)
                                elif file_name.endswith('.png'):
                                    preview_files.append(file_name)
                                elif 'metadata' in file_name.lower():
                                    metadata_files.append(file_name)
                            except Exception:
                                pass
                        
                        # 对G-code文件和预览图按plate编号排序
                        def get_plate_number(file_name):
                            """从文件名中提取plate编号"""
                            try:
                                match = re.search(r'plate_([0-9]+)', file_name)
                                if match:
                                    return int(match.group(1))
                                return 0
                            except Exception:
                                return 0
                        
                        # 排序文件
                        try:
                            sorted_gcode_files = sorted(gcode_files, key=get_plate_number)
                            sorted_preview_files = sorted(preview_files, key=get_plate_number)
                        except Exception:
                            sorted_gcode_files = gcode_files
                            sorted_preview_files = preview_files
                        
                        # 检查是否有多个打印盘
                        if len(sorted_gcode_files) > 1 or len(sorted_preview_files) > 1:
                            # 按plate分组文件
                            plate_groups = {}
                            
                            # 分组G-code文件
                            for gcode_file in sorted_gcode_files:
                                try:
                                    plate_num = get_plate_number(gcode_file)
                                    if plate_num not in plate_groups:
                                        plate_groups[plate_num] = {'gcode': [], 'preview': [], 'metadata': []}
                                    plate_groups[plate_num]['gcode'].append(gcode_file)
                                except Exception:
                                    pass
                            
                            # 分组预览图
                            for preview_file in sorted_preview_files:
                                try:
                                    plate_num = get_plate_number(preview_file)
                                    if plate_num not in plate_groups:
                                        plate_groups[plate_num] = {'gcode': [], 'preview': [], 'metadata': []}
                                    plate_groups[plate_num]['preview'].append(preview_file)
                                except Exception:
                                    pass
                            
                            # 分组元数据文件
                            for metadata_file in metadata_files:
                                try:
                                    plate_num = get_plate_number(metadata_file)
                                    if plate_num not in plate_groups:
                                        plate_groups[plate_num] = {'gcode': [], 'preview': [], 'metadata': []}
                                    plate_groups[plate_num]['metadata'].append(metadata_file)
                                except Exception:
                                    pass
                            
                            # 处理每个打印盘
                            try:
                                sorted_plate_nums = sorted(plate_groups.keys())
                            except Exception:
                                sorted_plate_nums = list(plate_groups.keys())
                            
                            for plate_num in sorted_plate_nums:
                                try:
                                    plate_info = "3MF文件信息\n"
                                    plate_info += "=" * 50 + "\n\n"
                                    plate_info += "文件路径\n"
                                    plate_info += "-" * 30 + "\n"
                                    plate_info += f"{file_path}\n"
                                    plate_info += f"打印盘: {plate_num}\n"
                                    
                                    # 添加文件大小
                                    plate_info += "\n文件大小\n"
                                    plate_info += "-" * 30 + "\n"
                                    try:
                                        file_size_bytes = os.path.getsize(file_path)
                                        file_size_mb = file_size_bytes / (1024 * 1024)
                                        plate_info += f"{file_size_bytes} 字节 ({file_size_mb:.2f} MB)\n"
                                    except Exception:
                                        plate_info += "无法获取文件大小\n"
                                    
                                    # 添加文件修改时间
                                    plate_info += "\n文件修改时间\n"
                                    plate_info += "-" * 30 + "\n"
                                    try:
                                        import datetime
                                        mod_time = os.path.getmtime(file_path)
                                        mod_time_str = datetime.datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                                        plate_info += f"修改时间: {mod_time_str}\n"
                                    except Exception:
                                        plate_info += "无法获取文件修改时间\n"
                                    
                                    plate_files = plate_groups[plate_num]
                                    plate_gcode_files = plate_files['gcode']
                                    plate_preview_files = plate_files['preview']
                                    plate_metadata_files = plate_files['metadata']
                                    
                                    # 读取G-code内容
                                    plate_gcode_content = ""
                                    for gcode_file in plate_gcode_files:
                                        try:
                                            # 检查文件大小，避免读取过大的文件
                                            gcode_file_info = zf.getinfo(gcode_file)
                                            file_size = gcode_file_info.file_size
                                            
                                            # 限制单个G-code文件最大为50MB
                                            if file_size > 50 * 1024 * 1024:
                                                plate_info += f"警告: G-code文件过大 ({file_size / 1024 / 1024:.2f}MB)，已跳过\n"
                                                continue
                                            
                                            with zf.open(gcode_file, 'r') as f:
                                                content = f.read().decode('utf-8', errors='ignore')
                                                plate_gcode_content += content
                                        except Exception as e:
                                            plate_info += f"读取G-code文件错误: {str(e)}\n"
                                    
                                    # 读取元数据
                                    plate_metadata = {}
                                    for metadata_file in plate_metadata_files:
                                        try:
                                            with zf.open(metadata_file, 'r') as f:
                                                content = f.read().decode('utf-8', errors='ignore')
                                                # 简单解析元数据
                                                for line in content.split('\n'):
                                                    try:
                                                        line = line.strip()
                                                        if ':' in line:
                                                            key, value = line.split(':', 1)
                                                            plate_metadata[key.strip()] = value.strip()
                                                    except Exception:
                                                        pass
                                        except Exception as e:
                                            plate_info += f"读取元数据文件错误: {str(e)}\n"
                                    
                                    # 提取预览图
                                    plate_preview_image_path = ""
                                    
                                    # 按plate编号分组预览图
                                    plate_preview_groups = {}
                                    for preview_file in plate_preview_files:
                                        try:
                                            if f'plate_{plate_num}' in preview_file:
                                                # 提取base名称（不含分辨率等后缀）
                                                base_name = re.sub(r'_\d+x\d+', '', preview_file)
                                                if base_name not in plate_preview_groups:
                                                    plate_preview_groups[base_name] = []
                                                plate_preview_groups[base_name].append(preview_file)
                                        except Exception:
                                            pass
                                    
                                    # 选择合适的预览图
                                    selected_preview = None
                                    for base_name, previews in plate_preview_groups.items():
                                        try:
                                            # 优先选择非"no light"的预览图
                                            no_light_previews = [p for p in previews if 'no light' in p.lower()]
                                            normal_previews = [p for p in previews if 'no light' not in p.lower()]
                                            
                                            if normal_previews:
                                                # 选择正常预览图中分辨率最高的
                                                normal_previews.sort(key=lambda x: int(re.search(r'(\d+)x(\d+)', x).group(1)) if re.search(r'(\d+)x(\d+)', x) else 0, reverse=True)
                                                selected_preview = normal_previews[0]
                                                break
                                            elif no_light_previews:
                                                # 如果只有"no light"预览图，选择分辨率最高的
                                                no_light_previews.sort(key=lambda x: int(re.search(r'(\d+)x(\d+)', x).group(1)) if re.search(r'(\d+)x(\d+)', x) else 0, reverse=True)
                                                selected_preview = no_light_previews[0]
                                                break
                                        except Exception:
                                            pass
                                    
                                    if selected_preview:
                                        try:
                                            # 提取预览图到临时文件
                                            with zf.open(selected_preview, 'r') as f:
                                                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                                                    temp_file.write(f.read())
                                                    plate_preview_image_path = temp_file.name
                                            plate_info += "\n预览图\n"
                                            plate_info += "-" * 30 + "\n"
                                            plate_info += f"{os.path.basename(selected_preview)}\n"
                                        except Exception as e:
                                            plate_info += f"提取预览图错误: {str(e)}\n"
                                    
                                    # 检查是否有G-code数据
                                    has_gcode_data = bool(plate_gcode_content)
                                    
                                    # 添加预览图信息
                                    if selected_preview:
                                        plate_info += "\n预览图信息\n"
                                        plate_info += "-" * 30 + "\n"
                                        plate_info += f"Plate {plate_num}: {os.path.basename(selected_preview)}\n"
                                    
                                    # 从G-code中提取打印时间和总估计时间
                                    print_time = ""
                                    total_time = ""
                                    
                                    if plate_gcode_content:
                                        # 查找所有打印时间（支持多种格式）
                                        print_time_matches = []
                                        
                                        # 格式1: model printing time: 7h 22m 20s
                                        matches = re.findall(r'model printing time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                        print_time_matches.extend(matches)
                                        
                                        # 格式2: printing time: 7h 22m 20s
                                        if not print_time_matches:
                                            matches = re.findall(r'printing time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            print_time_matches.extend(matches)
                                        
                                        # 格式3: model printing time: 35m 20s（小于一小时）
                                        if not print_time_matches:
                                            matches = re.findall(r'model printing time:\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            # 转换为 (0, minutes, seconds) 格式
                                            print_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                        
                                        # 格式4: printing time: 35m 20s（小于一小时）
                                        if not print_time_matches:
                                            matches = re.findall(r'printing time:\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            # 转换为 (0, minutes, seconds) 格式
                                            print_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                        
                                        # 计算总打印时间
                                        total_print_seconds = 0
                                        for match in print_time_matches:
                                            hours = int(match[0])
                                            minutes = int(match[1])
                                            seconds = int(match[2])
                                            total_print_seconds += hours * 3600 + minutes * 60 + seconds
                                        
                                        if total_print_seconds > 0:
                                            hours = total_print_seconds // 3600
                                            minutes = (total_print_seconds % 3600) // 60
                                            seconds = total_print_seconds % 60
                                            if total_print_seconds < 3600:
                                                # 少于一小时，显示 MM:SS
                                                print_time = f"{minutes:02d}:{seconds:02d}"
                                            else:
                                                # 一小时或以上，显示 HH:MM:SS
                                                print_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                                        
                                        # 查找所有总估计时间（支持多种格式）
                                        total_time_matches = []
                                        
                                        # 格式1: total estimated time: 7h 29m 48s
                                        matches = re.findall(r'total estimated time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                        total_time_matches.extend(matches)
                                        
                                        # 格式2: estimated time: 7h 29m 48s
                                        if not total_time_matches:
                                            matches = re.findall(r'estimated time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            total_time_matches.extend(matches)
                                        
                                        # 格式3: total estimated time: 35m 20s（小于一小时）
                                        if not total_time_matches:
                                            matches = re.findall(r'total estimated time:\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            # 转换为 (0, minutes, seconds) 格式
                                            total_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                        
                                        # 格式4: estimated time: 35m 20s（小于一小时）
                                        if not total_time_matches:
                                            matches = re.findall(r'estimated time:\s*(\d+)m\s*(\d+)s', plate_gcode_content)
                                            # 转换为 (0, minutes, seconds) 格式
                                            total_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                        
                                        # 计算总估计时间
                                        total_est_seconds = 0
                                        for match in total_time_matches:
                                            hours = int(match[0])
                                            minutes = int(match[1])
                                            seconds = int(match[2])
                                            total_est_seconds += hours * 3600 + minutes * 60 + seconds
                                        
                                        if total_est_seconds > 0:
                                            hours = total_est_seconds // 3600
                                            minutes = (total_est_seconds % 3600) // 60
                                            seconds = total_est_seconds % 60
                                            if total_est_seconds < 3600:
                                                # 少于一小时，显示 MM:SS
                                                total_time = f"{minutes:02d}:{seconds:02d}"
                                            else:
                                                # 一小时或以上，显示 HH:MM:SS
                                                total_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                                    
                                    # 添加打印时间信息
                                    if print_time:
                                        plate_info += f"\nPlate {plate_num} 打印时间\n"
                                        plate_info += "-" * 30 + "\n"
                                        plate_info += f"{print_time}\n"
                                    else:
                                        # 如果没有找到打印时间，添加默认值
                                        plate_info += f"\nPlate {plate_num} 打印时间\n"
                                        plate_info += "-" * 30 + "\n"
                                        plate_info += "00:00:00\n"
                                    
                                    # 添加总估计时间信息
                                    if total_time:
                                        plate_info += f"\nPlate {plate_num} 总估计时间\n"
                                        plate_info += "-" * 30 + "\n"
                                        plate_info += f"{total_time}\n"
                                    else:
                                        # 如果没有找到总估计时间，添加默认值
                                        plate_info += f"\nPlate {plate_num} 总估计时间\n"
                                        plate_info += "-" * 30 + "\n"
                                        plate_info += "00:00:00\n"
                                    
                                    # 只存储有G-code数据的打印盘
                                    if has_gcode_data:
                                        try:
                                            plate_data.append((file_path, plate_info, plate_preview_image_path, plate_gcode_content, plate_metadata.copy(), ""))
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                        else:
                            # 单打印盘处理
                            file_info = "3MF文件信息\n"
                            file_info += "=" * 50 + "\n\n"
                            file_info += "文件路径\n"
                            file_info += "-" * 30 + "\n"
                            file_info += f"{file_path}\n"
                            
                            # 添加文件大小
                            file_info += "\n文件大小\n"
                            file_info += "-" * 30 + "\n"
                            try:
                                file_size_bytes = os.path.getsize(file_path)
                                file_size_mb = file_size_bytes / (1024 * 1024)
                                file_info += f"{file_size_bytes} 字节 ({file_size_mb:.2f} MB)\n"
                            except Exception:
                                file_info += "无法获取文件大小\n"
                            
                            # 添加文件修改时间
                            file_info += "\n文件修改时间\n"
                            file_info += "-" * 30 + "\n"
                            try:
                                import datetime
                                mod_time = os.path.getmtime(file_path)
                                mod_time_str = datetime.datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d %H:%M:%S')
                                file_info += f"修改时间: {mod_time_str}\n"
                            except Exception:
                                file_info += "无法获取文件修改时间\n"
                            
                            # 读取G-code内容
                            gcode_content = ""
                            for gcode_file in gcode_files:
                                try:
                                    # 检查文件大小，避免读取过大的文件
                                    gcode_file_info = zf.getinfo(gcode_file)
                                    file_size = gcode_file_info.file_size
                                    
                                    # 限制单个G-code文件最大为50MB
                                    if file_size > 50 * 1024 * 1024:
                                        file_info += f"警告: G-code文件过大 ({file_size / 1024 / 1024:.2f}MB)，已跳过\n"
                                        continue
                                    
                                    with zf.open(gcode_file, 'r') as f:
                                        content = f.read().decode('utf-8', errors='ignore')
                                        gcode_content += content
                                except Exception as e:
                                    file_info += f"读取G-code文件错误: {str(e)}\n"
                            
                            # 读取元数据
                            metadata = {}
                            for metadata_file in metadata_files:
                                try:
                                    with zf.open(metadata_file, 'r') as f:
                                        content = f.read().decode('utf-8', errors='ignore')
                                        # 简单解析元数据
                                        for line in content.split('\n'):
                                            try:
                                                line = line.strip()
                                                if ':' in line:
                                                    key, value = line.split(':', 1)
                                                    metadata[key.strip()] = value.strip()
                                            except Exception:
                                                pass
                                except Exception as e:
                                    file_info += f"读取元数据文件错误: {str(e)}\n"
                            
                            # 提取预览图
                            preview_image_path = ""
                            if preview_files:
                                # 按base名称分组预览图
                                preview_groups = {}
                                for preview_file in preview_files:
                                    try:
                                        # 提取base名称（不含分辨率等后缀）
                                        base_name = re.sub(r'_\d+x\d+', '', preview_file)
                                        if base_name not in preview_groups:
                                            preview_groups[base_name] = []
                                        preview_groups[base_name].append(preview_file)
                                    except Exception:
                                        pass
                                
                                # 选择合适的预览图
                                selected_preview = None
                                for base_name, previews in preview_groups.items():
                                    try:
                                        # 优先选择非"no light"的预览图
                                        no_light_previews = [p for p in previews if 'no light' in p.lower()]
                                        normal_previews = [p for p in previews if 'no light' not in p.lower()]
                                        
                                        if normal_previews:
                                            # 选择正常预览图中分辨率最高的
                                            normal_previews.sort(key=lambda x: int(re.search(r'(\d+)x(\d+)', x).group(1)) if re.search(r'(\d+)x(\d+)', x) else 0, reverse=True)
                                            selected_preview = normal_previews[0]
                                            break
                                        elif no_light_previews:
                                            # 如果只有"no light"预览图，选择分辨率最高的
                                            no_light_previews.sort(key=lambda x: int(re.search(r'(\d+)x(\d+)', x).group(1)) if re.search(r'(\d+)x(\d+)', x) else 0, reverse=True)
                                            selected_preview = no_light_previews[0]
                                            break
                                    except Exception:
                                        pass
                                
                                if selected_preview:
                                    try:
                                        # 提取预览图到临时文件
                                        with zf.open(selected_preview, 'r') as f:
                                            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                                                temp_file.write(f.read())
                                                preview_image_path = temp_file.name
                                        file_info += "\n预览图信息\n"
                                        file_info += "-" * 30 + "\n"
                                        file_info += f"Plate 1: {os.path.basename(selected_preview)}\n"
                                    except Exception as e:
                                        file_info += f"提取预览图错误: {str(e)}\n"
                            
                            # 从G-code中提取打印时间和总估计时间
                            print_time = ""
                            total_time = ""
                            
                            if gcode_content:
                                # 查找所有打印时间（支持多种格式）
                                print_time_matches = []
                                
                                # 格式1: model printing time: 7h 22m 20s
                                matches = re.findall(r'model printing time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', gcode_content)
                                print_time_matches.extend(matches)
                                
                                # 格式2: printing time: 7h 22m 20s
                                if not print_time_matches:
                                    matches = re.findall(r'printing time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', gcode_content)
                                    print_time_matches.extend(matches)
                                
                                # 格式3: model printing time: 35m 20s（小于一小时）
                                if not print_time_matches:
                                    matches = re.findall(r'model printing time:\s*(\d+)m\s*(\d+)s', gcode_content)
                                    # 转换为 (0, minutes, seconds) 格式
                                    print_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                
                                # 格式4: printing time: 35m 20s（小于一小时）
                                if not print_time_matches:
                                    matches = re.findall(r'printing time:\s*(\d+)m\s*(\d+)s', gcode_content)
                                    # 转换为 (0, minutes, seconds) 格式
                                    print_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                
                                # 计算总打印时间
                                total_print_seconds = 0
                                for match in print_time_matches:
                                    hours = int(match[0])
                                    minutes = int(match[1])
                                    seconds = int(match[2])
                                    total_print_seconds += hours * 3600 + minutes * 60 + seconds
                                
                                if total_print_seconds > 0:
                                    hours = total_print_seconds // 3600
                                    minutes = (total_print_seconds % 3600) // 60
                                    seconds = total_print_seconds % 60
                                    if total_print_seconds < 3600:
                                        # 少于一小时，显示 MM:SS
                                        print_time = f"{minutes:02d}:{seconds:02d}"
                                    else:
                                        # 一小时或以上，显示 HH:MM:SS
                                        print_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                                
                                # 查找所有总估计时间（支持多种格式）
                                total_time_matches = []
                                
                                # 格式1: total estimated time: 7h 29m 48s
                                matches = re.findall(r'total estimated time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', gcode_content)
                                total_time_matches.extend(matches)
                                
                                # 格式2: estimated time: 7h 29m 48s
                                if not total_time_matches:
                                    matches = re.findall(r'estimated time:\s*(\d+)h\s*(\d+)m\s*(\d+)s', gcode_content)
                                    total_time_matches.extend(matches)
                                
                                # 格式3: total estimated time: 35m 20s（小于一小时）
                                if not total_time_matches:
                                    matches = re.findall(r'total estimated time:\s*(\d+)m\s*(\d+)s', gcode_content)
                                    # 转换为 (0, minutes, seconds) 格式
                                    total_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                
                                # 格式4: estimated time: 35m 20s（小于一小时）
                                if not total_time_matches:
                                    matches = re.findall(r'estimated time:\s*(\d+)m\s*(\d+)s', gcode_content)
                                    # 转换为 (0, minutes, seconds) 格式
                                    total_time_matches.extend([(0, match[0], match[1]) for match in matches])
                                
                                # 计算总估计时间
                                total_est_seconds = 0
                                for match in total_time_matches:
                                    hours = int(match[0])
                                    minutes = int(match[1])
                                    seconds = int(match[2])
                                    total_est_seconds += hours * 3600 + minutes * 60 + seconds
                                
                                if total_est_seconds > 0:
                                    hours = total_est_seconds // 3600
                                    minutes = (total_est_seconds % 3600) // 60
                                    seconds = total_est_seconds % 60
                                    if total_est_seconds < 3600:
                                        # 少于一小时，显示 MM:SS
                                        total_time = f"{minutes:02d}:{seconds:02d}"
                                    else:
                                        # 一小时或以上，显示 HH:MM:SS
                                        total_time = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                            
                            # 添加打印时间信息
                            if print_time:
                                file_info += f"\nPlate 1 打印时间\n"
                                file_info += "-" * 30 + "\n"
                                file_info += f"{print_time}\n"
                            else:
                                # 如果没有找到打印时间，添加默认值
                                file_info += f"\nPlate 1 打印时间\n"
                                file_info += "-" * 30 + "\n"
                                file_info += "00:00:00\n"
                            
                            # 添加总估计时间信息
                            if total_time:
                                file_info += f"\nPlate 1 总估计时间\n"
                                file_info += "-" * 30 + "\n"
                                file_info += f"{total_time}\n"
                            else:
                                # 如果没有找到总估计时间，添加默认值
                                file_info += f"\nPlate 1 总估计时间\n"
                                file_info += "-" * 30 + "\n"
                                file_info += "00:00:00\n"
                            
                            # 如果没有找到打印盘数据，返回单个文件数据
                            if not plate_data:
                                try:
                                    plate_data.append((file_path, file_info, preview_image_path, gcode_content, metadata.copy(), ""))
                                except Exception:
                                    pass
                    except Exception:
                        pass
            except Exception:
                pass
            
            return plate_data
        except Exception:
            return []
    
    def display_preview_image(self):
        """显示预览图"""
        if hasattr(self, 'preview_label') and self.preview_image_path and os.path.exists(self.preview_image_path):
            try:
                pixmap = QPixmap(self.preview_image_path)
                if not pixmap.isNull():
                    # 调整预览图大小以适应标签
                    scaled_pixmap = pixmap.scaled(
                        self.preview_label.size(),
                        Qt.KeepAspectRatio,
                        Qt.SmoothTransformation
                    )
                    self.preview_label.setPixmap(scaled_pixmap)
                else:
                    self.preview_label.setText("无法加载预览图")
            except Exception as e:
                self.preview_label.setText(f"加载预览图错误: {str(e)}")
        else:
            self.preview_label.setText("无预览图")
    
    def display_file_info(self):
        """显示文件信息"""
        if hasattr(self, 'info_text'):
            # 基础文件信息
            display_text = self.file_info
            
            # 添加换色代码分析结果（如果有）
            if hasattr(self, 'color_change_analysis'):
                analysis = self.color_change_analysis
                if analysis:
                    display_text += "\n\n换色代码分析\n"
                    display_text += "=" * 50 + "\n"
                    
                    if analysis.get('all_same', False):
                        display_text += "状态: 换色代码全部相同\n"
                        display_text += "提示: 会使用同一个颜色打印\n"
                        if analysis.get('codes'):
                            display_text += f"代码内容: {analysis['codes'][0]}\n"
                    else:
                        display_text += "状态: 存在换色\n"
                        display_text += "提示: 会使用不同颜色打印\n"
                        if analysis.get('codes'):
                            display_text += "\n各换色代码内容:\n"
                            for i, code in enumerate(analysis['codes']):
                                display_text += f"{i+1}. {code}\n"
            
            self.info_text.setPlainText(display_text)
    
    def display_error(self):
        """显示错误信息"""
        if hasattr(self, 'info_text'):
            self.info_text.setPlainText("无法解析3MF文件，请检查文件格式是否正确。")
        self.preview_label.setText("无预览图")
    
    def show_gcode(self):
        """显示G-code内容"""
        if self.gcode_content:
            # 创建临时文件保存G-code内容
            with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False) as temp_file:
                temp_file.write(self.gcode_content.encode('utf-8'))
                temp_file_path = temp_file.name
            
            # 尝试使用默认程序打开G-code文件
            try:
                os.startfile(temp_file_path)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法打开G-code文件: {str(e)}")
        else:
            QMessageBox.information(self, "提示", "当前文件没有G-code内容！")
    
    def add_gcode(self):
        """添加自定义G-code代码"""
        if not self.file_data or self.current_file_index < 0:
            QMessageBox.information(self, "提示", "请先打开一个文件！")
            return
        
        # 获取当前文件的自定义G-code
        current_custom_gcode = ""
        if 0 <= self.current_file_index < len(self.file_data):
            current_custom_gcode = self.file_data[self.current_file_index][5]
        
        # 创建默认的自定义G-code模板
        default_gcode = current_custom_gcode or """; 等待5分钟再执行换盘操作
G4 P300000

;========开始换盘 =================
G91;
G380 S3 Z-20 F1200
G380 S2 Z75 F1200
G380 S3 Z-20 F1200
G380 S2 Z75 F1200
G380 S3 Z-20 F1200
G380 S2 Z75 F1200
G380 S3 Z-20 F1200
G1 Z5 F1200
G90;
G28 Y;
G90;
G1 Y266 F2000;
G4 P1000
G91;
G380 S2 Z30 F1200
G90;
M211 Y0 Z0 ;
G91;
G90;
G1 Y50 F1000
G1 Y0 F2500
G91;
G380 S3 Z-20 F1200
G90;
G1 Y266 F2000
G1 Y43 F2000
G1 Y266 F2000
G1 Y43 F5000
G1 Y266 F2000
G1 Y-2 F7000
G1 Y150 F2000
;=======换板结束====================

G91; 切换到相对坐标模式
G380 S3 Z-50 F1200; Z轴向下移动50单位
G90; 切换回绝对坐标模式"""
        
        # 创建自定义对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("添加自定义G-code")
        dialog.setGeometry(100, 100, 600, 400)
        
        # 创建布局
        layout = QVBoxLayout(dialog)
        
        # 添加标签
        label = QLabel("请输入要添加的G-code:")
        layout.addWidget(label)
        
        # 添加文本框
        text_edit = QTextEdit()
        text_edit.setPlainText(default_gcode)
        layout.addWidget(text_edit)
        
        # 创建按钮布局
        button_layout = QHBoxLayout()
        
        # 添加"添加"按钮
        add_button = QPushButton("添加")
        add_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        
        # 添加"添加gcode到所有导入文件"按钮
        add_all_button = QPushButton("添加gcode到所有导入文件")
        add_all_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px;")
        
        # 添加"取消"按钮
        cancel_button = QPushButton("取消")
        cancel_button.setStyleSheet("background-color: #f44336; color: white; padding: 10px;")
        
        # 添加按钮到布局
        button_layout.addWidget(add_button)
        button_layout.addWidget(add_all_button)
        button_layout.addWidget(cancel_button)
        
        # 添加按钮布局到主布局
        layout.addLayout(button_layout)
        
        # 定义按钮点击事件
        result = {'action': 'cancel', 'gcode': ''}
        
        def on_add():
            result['action'] = 'add'
            result['gcode'] = text_edit.toPlainText()
            dialog.accept()
        
        def on_add_all():
            result['action'] = 'add_all'
            result['gcode'] = text_edit.toPlainText()
            dialog.accept()
        
        def on_cancel():
            result['action'] = 'cancel'
            dialog.reject()
        
        # 连接信号
        add_button.clicked.connect(on_add)
        add_all_button.clicked.connect(on_add_all)
        cancel_button.clicked.connect(on_cancel)
        
        # 显示对话框
        if dialog.exec_() == QDialog.Accepted:
            custom_gcode = result['gcode']
            action = result['action']
            
            if action == 'add':
                # 更新当前文件的自定义G-code
                if 0 <= self.current_file_index < len(self.file_data):
                    file_path, file_info, preview_image_path, gcode_content, metadata, _ = self.file_data[self.current_file_index]
                    self.file_data[self.current_file_index] = (
                        file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode
                    )
                    self.custom_gcode = custom_gcode
                    QMessageBox.information(self, "成功", "自定义G-code已添加到当前文件！")
                else:
                    QMessageBox.critical(self, "错误", "无法更新自定义G-code，请先选择一个文件！")
            
            elif action == 'add_all':
                # 更新所有文件的自定义G-code
                if self.file_data:
                    for i in range(len(self.file_data)):
                        file_path, file_info, preview_image_path, gcode_content, metadata, _ = self.file_data[i]
                        self.file_data[i] = (
                            file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode
                        )
                    # 更新当前文件的custom_gcode
                    if 0 <= self.current_file_index < len(self.file_data):
                        self.custom_gcode = custom_gcode
                    QMessageBox.information(self, "成功", "自定义G-code已添加到所有导入文件！")
                else:
                    QMessageBox.critical(self, "错误", "没有导入任何文件！")
    
    def search_parameters(self):
        """搜索参数"""
        if not self.gcode_content:
            QMessageBox.information(self, "提示", "当前文件没有G-code内容！")
            return
        
        # 搜索层高等参数
        layer_pattern = r'LAYER:(\d+)'
        layer_matches = re.findall(layer_pattern, self.gcode_content)
        
        if layer_matches:
            max_layer = max(map(int, layer_matches))
            QMessageBox.information(
                self, "参数信息", f"最大层数: {max_layer}\n总层数: {len(layer_matches)}"
            )
        else:
            QMessageBox.information(self, "提示", "未找到层数信息！")
    
    def compare_color_change_codes(self):
        """换色代码差异对比"""
        if not self.gcode_content:
            QMessageBox.information(self, "提示", "当前文件没有G-code内容！")
            return
        
        # 提取所有M621代码行
        import re
        lines = self.gcode_content.split('\n')
        m621_lines = []
        
        for line in lines:
            line = line.strip()
            if line.startswith('M621'):
                # 提取完整的M621行
                m621_lines.append(line)
        
        # 分析差异
        if not m621_lines:
            QMessageBox.information(self, "分析结果", "未找到换色代码(M621)！")
            return
        
        # 检查是否所有行都相同
        all_same = all(line == m621_lines[0] for line in m621_lines)
        
        # 构建分析结果
        result = "换色代码分析结果:\n\n"
        if all_same:
            result += "状态: 换色代码全部相同\n"
            result += "提示: 会使用同一个颜色打印\n"
            result += f"代码内容: {m621_lines[0]}\n"
        else:
            result += "状态: 存在换色\n"
            result += "提示: 会使用不同颜色打印\n"
            result += "\n各换色代码内容:\n"
            for i, line in enumerate(m621_lines):
                result += f"{i+1}. {line}\n"
        
        # 显示分析结果
        QMessageBox.information(self, "换色代码分析", result)
        
        # 将分析结果添加到元数据中，以便在界面上显示
        if not hasattr(self, 'color_change_analysis'):
            self.color_change_analysis = {}
        
        self.color_change_analysis['all_same'] = all_same
        self.color_change_analysis['codes'] = m621_lines
        self.color_change_analysis['result'] = result
        
        # 重新显示文件信息，以便包含新的分析结果
        self.display_file_info()
    
    def export_file(self):
        """导出3MF文件"""
        if not self.file_data:
            QMessageBox.information(self, "提示", "请先打开一个文件！")
            return
        
        # 获取原文件名
        if self.file_data and len(self.file_data) > 0:
            file_path = self.file_data[0][0]
            # 提取原文件名（去掉.gcode.3mf后缀）
            base_name = os.path.basename(file_path)
            # 处理.gcode.3mf后缀
            if base_name.endswith('.gcode.3mf'):
                base_name = base_name[:-11]  # 去掉.gcode.3mf
            elif base_name.endswith('.3mf'):
                base_name = base_name[:-4]  # 去掉.3mf
            # 构建默认文件名：原文件名+最终.3mf
            default_filename = f"{base_name}最终.3mf"
        else:
            default_filename = "最终.3mf"
        
        # 获取导出文件路径
        export_path, _ = QFileDialog.getSaveFileName(
            self, "导出3MF文件", default_filename, "3MF Files (*.3mf);;All Files (*.*)"
        )
        
        if export_path:
            try:
                # 读取原始3MF文件
                original_file = self.file_data[0][0]
                with zipfile.ZipFile(original_file, 'r') as zf:
                    # 创建新的3MF文件
                    with zipfile.ZipFile(export_path, 'w', zipfile.ZIP_DEFLATED) as new_zf:
                        # 只保留必要的文件，去掉冗余内容
                        for file_info in zf.filelist:
                            filename = file_info.filename
                            
                            # 只保留G-code文件和元数据文件，去掉预览图等冗余文件
                            if filename.endswith('.gcode') or filename.endswith('.gcode.3mf'):
                                # 读取G-code文件内容
                                content = zf.read(filename).decode('utf-8', errors='ignore')
                                
                                # 添加自定义G-code
                                if self.file_data and len(self.file_data) > 0:
                                    custom_gcode = self.file_data[0][6]
                                    if custom_gcode:
                                        # 在G-code末尾添加自定义G-code
                                        content += f"\n; ====== 自定义G-code ======\n{custom_gcode}"
                                
                                # 写入新的G-code内容
                                new_zf.writestr(filename, content.encode('utf-8'))
                            elif filename.endswith('.metadata'):
                                # 保留元数据文件
                                new_zf.writestr(filename, zf.read(filename))
                            elif filename.endswith('.model'):
                                # 保留模型文件
                                new_zf.writestr(filename, zf.read(filename))
                            elif filename.endswith('.rels'):
                                # 保留关系文件
                                new_zf.writestr(filename, zf.read(filename))
                
                QMessageBox.information(self, "成功", f"文件已导出到: {export_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出文件时发生错误: {str(e)}")
    
    def export_final_file(self):
        """导出最终的3MF文件（含有自定义修改G-code内容）"""
        if not self.file_data:
            QMessageBox.information(self, "提示", "请先打开一个文件！")
            return
        
        # 获取原文件名
        if self.file_data and len(self.file_data) > 0:
            file_path = self.file_data[0][0]
            # 提取原文件名（去掉.gcode.3mf后缀）
            base_name = os.path.basename(file_path)
            # 处理.gcode.3mf后缀
            if base_name.endswith('.gcode.3mf'):
                base_name = base_name[:-11]  # 去掉.gcode.3mf
            elif base_name.endswith('.3mf'):
                base_name = base_name[:-4]  # 去掉.3mf
            # 构建默认文件名：原文件名+最终.3mf
            default_filename = f"{base_name}最终.3mf"
        else:
            default_filename = "最终.3mf"
        
        # 获取导出文件路径
        export_path, _ = QFileDialog.getSaveFileName(
            self, "导出最终3MF文件", default_filename, "3MF Files (*.3mf);;All Files (*.*)"
        )
        
        if export_path:
            try:
                # 读取原始3MF文件
                original_file = self.file_data[0][0]
                with zipfile.ZipFile(original_file, 'r') as zf:
                    # 创建新的3MF文件
                    with zipfile.ZipFile(export_path, 'w', zipfile.ZIP_DEFLATED) as new_zf:
                        # 合并多个打印盘的G-code
                        merged_gcode = ""
                        
                        # 按顺序处理每个打印盘
                        for i, plate in enumerate(self.file_data):
                            file_path, file_info, preview_image_path, gcode_content, metadata, custom_gcode = plate
                            
                            # 添加打印盘信息
                            merged_gcode += f"\n; ====== 打印盘 {i+1} ======\n"
                            merged_gcode += f"; 文件: {os.path.basename(file_path)}\n"
                            
                            # 添加原始G-code
                            merged_gcode += gcode_content
                            
                            # 添加自定义G-code
                            if custom_gcode:
                                merged_gcode += f"\n; ====== 自定义G-code ======\n"
                                merged_gcode += custom_gcode
                        
                        # 首先收集所有G-code文件
                        gcode_files = []
                        for file_info in zf.filelist:
                            filename = file_info.filename
                            if filename.endswith('.gcode') or filename.endswith('.gcode.3mf'):
                                gcode_files.append(filename)
                        
                        # 只保留第一个G-code文件，其余的都不保留
                        first_gcode_file = None
                        if gcode_files:
                            first_gcode_file = gcode_files[0]
                        
                        # 处理所有文件
                        for file_info in zf.filelist:
                            filename = file_info.filename
                            
                            # 只修改第一个G-code文件，其余的文件保持不变
                            if filename.endswith('.gcode') or filename.endswith('.gcode.3mf'):
                                if filename == first_gcode_file:
                                    # 写入合并后的G-code内容到第一个G-code文件
                                    new_zf.writestr(filename, merged_gcode.encode('utf-8'))
                                # 不保留其他G-code文件
                            else:
                                # 保留所有其他文件，包括：
                                # - [Content_Types].xml
                                # - _rels/.rels
                                # - 3D/3dmodel.model
                                # - Metadata/ 文件夹中的文件
                                # - 所有预览图
                                # - 所有其他必要的文件
                                new_zf.writestr(filename, zf.read(filename))
                
                QMessageBox.information(self, "成功", f"最终3MF文件已导出到: {export_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出最终文件时发生错误: {str(e)}")
    
    def clear_files(self):
        """清空文件"""
        self.three_mf_files = []
        self.file_data = []
        self.file_selector.clear()
        self.file_info = ""
        self.preview_image_path = ""
        self.gcode_content = ""
        self.metadata = {}
        self.custom_gcode = ""
        self.current_file = ""
        self.current_file_index = -1
        
        if hasattr(self, 'info_text'):
            self.info_text.setPlainText("3MF文件预览工具\n\n点击'打开3MF文件...'按钮选择要预览的3MF文件。")
        if hasattr(self, 'preview_label'):
            self.preview_label.setText("无预览图")
    
    def about(self):
        """显示关于信息"""
        about_text = """
3MF文件预览工具
版本: v1.5

功能:
1. 打开和预览3MF文件
2. 查看文件中的G-code内容
3. 支持多打印盘3MF文件
4. 显示文件的元数据信息
5. 添加自定义G-code代码
6. 导出修改后的3MF文件
7. 导出最终的G-code文件
8. 换色代码差异对比

更新内容:
- 优化多打印盘识别和排序
- 改进预览图选择逻辑
- 隐藏G-code验证和元数据提示
- 添加自定义G-code模板
- 修复文件处理bug
- 新增换色代码差异对比功能

作者: 纸上滑雪
"""
        
        # 创建关于对话框
        about_dialog = QMessageBox(self)
        about_dialog.setWindowTitle("关于")
        about_dialog.setText(about_text)
        
        # 加载图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'figurine64.ico')
        if os.path.exists(icon_path):
            # 加载图标并设置到对话框
            icon = QIcon(icon_path)
            about_dialog.setIconPixmap(icon.pixmap(64, 64))
            about_dialog.setWindowIcon(icon)
        
        about_dialog.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ThreeMfPreviewer()
    window.show()
    sys.exit(app.exec_())
