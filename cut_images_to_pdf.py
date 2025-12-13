#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片切割并转换为PDF工具
将文件名格式为 xxx-N.png 的图片切割成N份，并生成PDF
使用 PyObjC 实现原生 macOS GUI
"""

import os
import re
import threading
from queue import Queue
from PIL import Image
from Foundation import NSObject, NSLog, NSTimer
from AppKit import (NSApplication, NSWindow, NSButton, NSTextField, NSTextView,
                    NSScrollView, NSBox, NSAlert, NSOpenPanel, NSApp,
                    NSApplicationActivationPolicyRegular, NSBackingStoreBuffered,
                    NSWindowStyleMaskTitled, NSWindowStyleMaskClosable,
                    NSWindowStyleMaskMiniaturizable, NSWindowStyleMaskResizable,
                    NSMakeRect, NSMakeSize, NSFont, NSColor, NSBezelBorder,
                    NSTextFieldSquareBezel, NSPushOnPushOffButton,
                    NSAlertFirstButtonReturn, NSAlertSecondButtonReturn,
                    NSInformationalAlertStyle, NSWarningAlertStyle)
from objc import super as objc_super


def extract_page_count(filename):
    """
    从文件名中提取页数
    例如: "原密协议-5.png" -> 5
    """
    match = re.search(r'-(\d+)\.(png|jpg|jpeg|PNG|JPG|JPEG)$', filename)
    if match:
        return int(match.group(1))
    return None


def get_base_filename(input_filename):
    """
    获取基础文件名（不含扩展名和页数）
    例如: "原密协议-5.png" -> "原密协议"
    """
    match = re.search(r'^(.+?)-\d+\.(png|jpg|jpeg|PNG|JPG|JPEG)$', input_filename)
    if match:
        return match.group(1)
    return None


def get_output_filename(input_filename):
    """
    生成输出PDF文件名
    例如: "原密协议-5.png" -> "原密协议.pdf"
    """
    return re.sub(r'-\d+\.(png|jpg|jpeg|PNG|JPG|JPEG)$', '.pdf', input_filename)


def cut_image_vertically(image_path, num_parts):
    """
    将图片垂直切割成num_parts份
    返回切割后的图片列表
    """
    img = Image.open(image_path)
    width, height = img.size
    
    # 计算每一部分的高度
    part_height = height // num_parts
    
    images = []
    for i in range(num_parts):
        # 计算切割区域
        top = i * part_height
        bottom = (i + 1) * part_height if i < num_parts - 1 else height
        
        # 切割图片
        box = (0, top, width, bottom)
        part_img = img.crop(box)
        images.append(part_img)
    
    img.close()
    return images


def save_cut_images(images, base_filename, output_folder):
    """
    保存切割后的图片
    """
    saved_paths = []
    for i, img in enumerate(images, 1):
        output_path = os.path.join(output_folder, f"{base_filename}-第{i}页.png")
        img.save(output_path, 'PNG')
        saved_paths.append(output_path)
    return saved_paths


def images_to_pdf(images, output_path):
    """
    将图片列表转换为PDF
    """
    if not images:
        return
    
    # 将所有图片转换为RGB模式（PDF需要）
    rgb_images = []
    for img in images:
        if img.mode != 'RGB':
            rgb_img = img.convert('RGB')
            rgb_images.append(rgb_img)
        else:
            rgb_images.append(img)
    
    # 保存为PDF
    if len(rgb_images) > 1:
        rgb_images[0].save(
            output_path,
            save_all=True,
            append_images=rgb_images[1:],
            resolution=100.0
        )
    else:
        rgb_images[0].save(output_path)
    
    # 清理
    for img in rgb_images:
        img.close()


class ImageCutterApp(NSObject):
    """图片切割工具的主应用类"""
    
    def init(self):
        self = objc_super(ImageCutterApp, self).init()
        if self is None:
            return None
        
        self.input_folder = ""
        self.output_folder = ""
        self.is_processing = False
        self.log_queue = Queue()  # 日志消息队列
        
        return self
    
    def applicationDidFinishLaunching_(self, notification):
        """应用启动完成"""
        self.create_window()
        # 启动定时器，定期检查日志队列
        self.timer = NSTimer.scheduledTimerWithTimeInterval_target_selector_userInfo_repeats_(
            0.1,  # 每100ms检查一次
            self,
            "updateLogFromQueue:",
            None,
            True
        )
    
    def create_window(self):
        """创建主窗口"""
        # 创建窗口
        frame = NSMakeRect(100, 100, 750, 550)
        style = (NSWindowStyleMaskTitled | NSWindowStyleMaskClosable |
                NSWindowStyleMaskMiniaturizable | NSWindowStyleMaskResizable)
        
        self.window = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(
            frame, style, NSBackingStoreBuffered, False
        )
        self.window.setTitle_("图片切割转PDF工具")
        self.window.setMinSize_(NSMakeSize(600, 400))
        self.window.center()
        
        # 创建UI元素
        y_pos = 510
        
        # 顶部说明标签
        info_label = NSTextField.alloc().initWithFrame_(NSMakeRect(20, y_pos, 710, 30))
        info_label.setStringValue_("将文件名格式为 xxx-N.png 的图片切割成N份，并生成PDF")
        info_label.setBezeled_(False)
        info_label.setDrawsBackground_(False)
        info_label.setEditable_(False)
        info_label.setSelectable_(False)
        info_label.setAlignment_(1)  # 居中
        info_label.setFont_(NSFont.systemFontOfSize_(13))
        info_label.setTextColor_(NSColor.secondaryLabelColor())
        self.window.contentView().addSubview_(info_label)
        
        y_pos -= 50
        
        # 输入文件夹组
        input_box = NSBox.alloc().initWithFrame_(NSMakeRect(20, y_pos, 710, 70))
        input_box.setTitle_("输入文件夹")
        input_box.setTitlePosition_(2)  # NSAtTop
        
        self.input_field = NSTextField.alloc().initWithFrame_(NSMakeRect(15, 15, 565, 24))
        self.input_field.setPlaceholderString_("选择包含图片的文件夹...")
        self.input_field.setEditable_(False)
        self.input_field.setSelectable_(True)
        self.input_field.setBezeled_(True)
        self.input_field.setBezelStyle_(NSTextFieldSquareBezel)
        self.input_field.setDrawsBackground_(True)
        self.input_field.setBackgroundColor_(NSColor.textBackgroundColor())
        input_box.contentView().addSubview_(self.input_field)
        
        input_btn = NSButton.alloc().initWithFrame_(NSMakeRect(590, 15, 100, 28))
        input_btn.setTitle_("选择")
        input_btn.setBezelStyle_(1)  # NSRoundedBezelStyle
        input_btn.setTarget_(self)
        input_btn.setAction_("selectInputFolder:")
        input_box.contentView().addSubview_(input_btn)
        
        self.window.contentView().addSubview_(input_box)
        
        y_pos -= 85
        
        # 输出文件夹组
        output_box = NSBox.alloc().initWithFrame_(NSMakeRect(20, y_pos, 710, 70))
        output_box.setTitle_("输出文件夹")
        output_box.setTitlePosition_(2)  # NSAtTop
        
        self.output_field = NSTextField.alloc().initWithFrame_(NSMakeRect(15, 15, 565, 24))
        self.output_field.setPlaceholderString_("选择输出文件夹（可选，默认同输入文件夹）...")
        self.output_field.setEditable_(False)
        self.output_field.setSelectable_(True)
        self.output_field.setBezeled_(True)
        self.output_field.setBezelStyle_(NSTextFieldSquareBezel)
        self.output_field.setDrawsBackground_(True)
        self.output_field.setBackgroundColor_(NSColor.textBackgroundColor())
        output_box.contentView().addSubview_(self.output_field)
        
        output_btn = NSButton.alloc().initWithFrame_(NSMakeRect(590, 15, 100, 28))
        output_btn.setTitle_("选择")
        output_btn.setBezelStyle_(1)
        output_btn.setTarget_(self)
        output_btn.setAction_("selectOutputFolder:")
        output_box.contentView().addSubview_(output_btn)
        
        self.window.contentView().addSubview_(output_box)
        
        y_pos -= 65
        
        # 开始处理按钮
        self.process_btn = NSButton.alloc().initWithFrame_(NSMakeRect(275, y_pos, 200, 38))
        self.process_btn.setTitle_("开始处理")
        self.process_btn.setBezelStyle_(1)
        self.process_btn.setFont_(NSFont.boldSystemFontOfSize_(15))
        self.process_btn.setTarget_(self)
        self.process_btn.setAction_("startProcessing:")
        self.window.contentView().addSubview_(self.process_btn)
        
        y_pos -= 55
        
        # 日志输出区域
        log_box = NSBox.alloc().initWithFrame_(NSMakeRect(20, 20, 710, y_pos - 20))
        log_box.setTitle_("处理日志")
        log_box.setTitlePosition_(2)  # NSAtTop
        
        # 计算scrollview的高度
        scroll_height = y_pos - 55
        scroll_view = NSScrollView.alloc().initWithFrame_(NSMakeRect(15, 15, 680, scroll_height))
        scroll_view.setHasVerticalScroller_(True)
        scroll_view.setHasHorizontalScroller_(False)
        scroll_view.setAutohidesScrollers_(True)
        scroll_view.setBorderType_(NSBezelBorder)
        
        text_frame = NSMakeRect(0, 0, 660, scroll_height)
        self.log_view = NSTextView.alloc().initWithFrame_(text_frame)
        self.log_view.setEditable_(False)
        self.log_view.setSelectable_(True)
        self.log_view.setFont_(NSFont.fontWithName_size_("Menlo", 11))
        self.log_view.setTextContainerInset_(NSMakeSize(5, 5))
        scroll_view.setDocumentView_(self.log_view)
        
        log_box.contentView().addSubview_(scroll_view)
        self.window.contentView().addSubview_(log_box)
        
        # 显示窗口
        self.window.makeKeyAndOrderFront_(None)
    
    def selectInputFolder_(self, sender):
        """选择输入文件夹"""
        panel = NSOpenPanel.openPanel()
        panel.setCanChooseFiles_(False)
        panel.setCanChooseDirectories_(True)
        panel.setAllowsMultipleSelection_(False)
        panel.setPrompt_("选择")
        panel.setMessage_("选择包含图片的文件夹")
        
        result = panel.runModal()
        if result == 1:  # NSModalResponseOK
            urls = panel.URLs()
            if urls and len(urls) > 0:
                url = urls[0]
                path = url.path()
                self.input_folder = path
                self.input_field.setStringValue_(path)
                self.input_field.setNeedsDisplay_(True)  # 强制刷新显示
                self.appendLog_(f"已选择输入文件夹: {path}")
                
                # 如果输出文件夹为空，默认使用输入文件夹
                if not self.output_folder:
                    self.output_folder = path
                    self.output_field.setStringValue_(path)
                    self.output_field.setNeedsDisplay_(True)
    
    def selectOutputFolder_(self, sender):
        """选择输出文件夹"""
        panel = NSOpenPanel.openPanel()
        panel.setCanChooseFiles_(False)
        panel.setCanChooseDirectories_(True)
        panel.setAllowsMultipleSelection_(False)
        panel.setPrompt_("选择")
        panel.setMessage_("选择输出文件夹")
        
        result = panel.runModal()
        if result == 1:  # NSModalResponseOK
            urls = panel.URLs()
            if urls and len(urls) > 0:
                url = urls[0]
                path = url.path()
                self.output_folder = path
                self.output_field.setStringValue_(path)
                self.output_field.setNeedsDisplay_(True)  # 强制刷新显示
                self.appendLog_(f"已选择输出文件夹: {path}")
    
    def appendLog_(self, message):
        """添加日志信息（线程安全）- 将消息放入队列"""
        self.log_queue.put(message)
    
    def updateLogFromQueue_(self, timer):
        """从队列中取出日志消息并更新UI（在主线程中调用）"""
        try:
            while not self.log_queue.empty():
                message = self.log_queue.get_nowait()
                current_text = self.log_view.string()
                new_text = current_text + message + "\n"
                self.log_view.setString_(new_text)
                # 滚动到底部
                self.log_view.scrollRangeToVisible_((len(new_text), 0))
        except:
            pass  # 队列为空或其他错误，忽略
    
    def startProcessing_(self, sender):
        """开始处理图片"""
        if self.is_processing:
            self.show_alert("警告", "正在处理中，请稍候...", NSWarningAlertStyle)
            return
        
        # 直接使用保存的变量而不是从TextField读取
        input_path = self.input_folder
        output_path = self.output_folder
        
        if not input_path:
            self.show_alert("错误", "请先选择输入文件夹！", NSWarningAlertStyle)
            return
        
        if not output_path:
            output_path = input_path
            self.output_folder = output_path
            self.output_field.setStringValue_(output_path)
        
        # 清空日志
        self.log_view.setString_("")
        self.appendLog_("=" * 60)
        self.appendLog_("开始处理...")
        self.appendLog_(f"输入文件夹: {input_path}")
        self.appendLog_(f"输出文件夹: {output_path}")
        self.appendLog_("=" * 60 + "\n")
        
        # 禁用按钮
        self.is_processing = True
        self.process_btn.setEnabled_(False)
        self.process_btn.setTitle_("处理中...")
        
        # 在新线程中处理
        def process_thread():
            try:
                success = self.process_folder(input_path, output_path)
                self.on_processing_finished(success)
            except Exception as e:
                self.appendLog_(f"\n错误: {str(e)}")
                self.on_processing_finished(False)
        
        thread = threading.Thread(target=process_thread, daemon=True)
        thread.start()
    
    def process_folder(self, folder_path, output_folder):
        """处理文件夹中的所有图片"""
        if not os.path.isdir(folder_path):
            self.appendLog_(f"错误: {folder_path} 不是一个有效的文件夹")
            return False
        
        # 确保输出文件夹存在
        os.makedirs(output_folder, exist_ok=True)
        
        # 获取所有图片文件
        files = os.listdir(folder_path)
        image_files = [f for f in files if re.search(r'-\d+\.(png|jpg|jpeg|PNG|JPG|JPEG)$', f)]
        
        if not image_files:
            self.appendLog_(f"在 {folder_path} 中没有找到符合格式的图片文件")
            self.appendLog_("文件名格式应为: xxx-N.png (N为页数)")
            return False
        
        self.appendLog_(f"找到 {len(image_files)} 个图片文件")
        
        success_count = 0
        error_count = 0
        
        for filename in image_files:
            try:
                # 提取页数
                page_count = extract_page_count(filename)
                if page_count is None:
                    self.appendLog_(f"跳过 {filename}: 无法提取页数")
                    continue
                
                self.appendLog_(f"\n处理 {filename} (共 {page_count} 页)...")
                
                # 完整路径
                input_path = os.path.join(folder_path, filename)
                
                # 获取基础文件名
                base_filename = get_base_filename(filename)
                if base_filename is None:
                    self.appendLog_(f"  跳过 {filename}: 无法提取基础文件名")
                    continue
                
                # 切割图片
                self.appendLog_(f"  正在切割图片...")
                images = cut_image_vertically(input_path, page_count)
                self.appendLog_(f"  已切割为 {len(images)} 部分")
                
                # 保存切割后的图片
                self.appendLog_(f"  正在保存切割后的图片...")
                saved_paths = save_cut_images(images, base_filename, output_folder)
                self.appendLog_(f"  已保存 {len(saved_paths)} 张图片")
                
                # 生成PDF
                output_filename = get_output_filename(filename)
                output_path = os.path.join(output_folder, output_filename)
                
                self.appendLog_(f"  正在生成PDF: {output_filename}...")
                images_to_pdf(images, output_path)
                
                self.appendLog_(f"  ✓ 完成: {output_filename}")
                success_count += 1
                
            except Exception as e:
                self.appendLog_(f"  ✗ 处理 {filename} 时出错: {str(e)}")
                error_count += 1
        
        self.appendLog_(f"\n所有文件处理完成！成功: {success_count}, 失败: {error_count}")
        return True
    
    def on_processing_finished(self, success):
        """处理完成回调 - 使用定时器调度UI更新"""
        self.finish_success = success
        # 使用单次定时器在主线程执行
        NSTimer.scheduledTimerWithTimeInterval_target_selector_userInfo_repeats_(
            0.01,
            self,
            "updateUIAfterProcessing:",
            None,
            False
        )
    
    def updateUIAfterProcessing_(self, timer):
        """在主线程中更新UI"""
        self.is_processing = False
        self.process_btn.setEnabled_(True)
        self.process_btn.setTitle_("开始处理")
        
        success = getattr(self, 'finish_success', False)
        if success:
            self.show_alert("成功", "所有文件处理完成！", NSInformationalAlertStyle)
        else:
            self.show_alert("提示", "处理过程中出现问题，请查看日志", NSWarningAlertStyle)
    
    def show_alert(self, title, message, style):
        """显示警告对话框"""
        alert = NSAlert.alloc().init()
        alert.setMessageText_(title)
        alert.setInformativeText_(message)
        alert.setAlertStyle_(style)
        alert.addButtonWithTitle_("确定")
        alert.runModal()


def main():
    """主函数"""
    # 创建应用
    app = NSApplication.sharedApplication()
    app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
    
    # 创建应用委托
    delegate = ImageCutterApp.alloc().init()
    app.setDelegate_(delegate)
    
    # 激活应用
    app.activateIgnoringOtherApps_(True)
    
    # 运行应用
    app.run()


if __name__ == '__main__':
    main()
