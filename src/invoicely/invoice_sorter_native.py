#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Invoicely 原生 macOS 界面

无浏览器、无本地端口。关闭窗口后 App 直接退出，不保留后台服务。
"""

from __future__ import annotations

import threading
import traceback
from datetime import datetime
from pathlib import Path

import objc
from AppKit import (
    NSAlert,
    NSApplication,
    NSApplicationActivationPolicyRegular,
    NSBackingStoreBuffered,
    NSBezelStyleRounded,
    NSButton,
    NSButtonTypeSwitch,
    NSColor,
    NSMakeRect,
    NSMenu,
    NSMenuItem,
    NSOpenPanel,
    NSScrollView,
    NSTabView,
    NSTabViewItem,
    NSTextField,
    NSTextView,
    NSWindow,
    NSWindowStyleMaskClosable,
    NSWindowStyleMaskMiniaturizable,
    NSWindowStyleMaskResizable,
    NSWindowStyleMaskTitled,
    NSWorkspace,
)
from Foundation import NSObject
from PyObjCTools import AppHelper

from invoicely.invoice_core import (
    AppSettings,
    RateManager,
    build_excel,
    iter_invoice_files,
    normalize_folder_input,
    process_one_file,
)


class InvoiceNativeApp(NSObject):
    def applicationDidFinishLaunching_(self, notification):
        self.excel_path = None
        self.processing = False
        self.settings_window = None
        self.settings = AppSettings.load()
        self.rates = RateManager()
        self._build_menu()
        self._build_window()
        if self.settings.auto_update_rates:
            threading.Thread(target=self._refresh_rates_on_launch, daemon=True).start()

    def applicationShouldTerminateAfterLastWindowClosed_(self, sender):
        return True

    @objc.python_method
    def _build_menu(self):
        main_menu = NSMenu.alloc().init()
        app_menu_item = NSMenuItem.alloc().init()
        main_menu.addItem_(app_menu_item)

        app_menu = NSMenu.alloc().initWithTitle_("Invoicely")
        settings_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_("设置...", "openSettings:", ",")
        settings_item.setTarget_(self)
        app_menu.addItem_(settings_item)
        app_menu.addItem_(NSMenuItem.separatorItem())

        quit_item = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_("退出 Invoicely", "terminate:", "q")
        app_menu.addItem_(quit_item)
        app_menu_item.setSubmenu_(app_menu)
        NSApplication.sharedApplication().setMainMenu_(main_menu)

    @objc.python_method
    def _build_window(self):
        style = (
            NSWindowStyleMaskTitled
            | NSWindowStyleMaskClosable
            | NSWindowStyleMaskMiniaturizable
            | NSWindowStyleMaskResizable
        )
        self.window = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(
            NSMakeRect(0, 0, 820, 620),
            style,
            NSBackingStoreBuffered,
            False,
        )
        self.window.setTitle_("Invoicely 发票整理")
        self.window.center()

        content = self.window.contentView()

        title = NSTextField.alloc().initWithFrame_(NSMakeRect(28, 570, 360, 30))
        title.setStringValue_("Invoicely")
        title.setEditable_(False)
        title.setBordered_(False)
        title.setDrawsBackground_(False)
        title.setFont_(title.font().fontWithSize_(22))
        content.addSubview_(title)

        self.rate_label = self._label(NSMakeRect(560, 575, 232, 18), self.rates.short_status())
        self.rate_label.setAlignment_(2)
        content.addSubview_(self.rate_label)

        self.tab_view = NSTabView.alloc().initWithFrame_(NSMakeRect(28, 474, 764, 78))
        self.tab_view.addTabViewItem_(self._mode_tab("Mainland", "Mainland", "统一统计为 RMB，外币票据会自动换算并标注。"))
        self.tab_view.addTabViewItem_(self._mode_tab("HK", "HK", "统一统计为 HKD，港币票据不标注，其他币种换算后标注。"))
        content.addSubview_(self.tab_view)

        self.path_field = NSTextField.alloc().initWithFrame_(NSMakeRect(28, 428, 606, 30))
        self.path_field.setPlaceholderString_("把文件夹拖到这里，或点击右侧选择")
        content.addSubview_(self.path_field)

        choose_button = self._button(NSMakeRect(648, 427, 144, 32), "选择文件夹", "chooseFolder:")
        content.addSubview_(choose_button)

        self.start_button = self._button(NSMakeRect(28, 382, 126, 34), "开始处理", "startProcessing:")
        content.addSubview_(self.start_button)

        self.open_excel_button = self._button(NSMakeRect(170, 382, 140, 34), "打开统计表", "openExcel:")
        self.open_excel_button.setEnabled_(False)
        content.addSubview_(self.open_excel_button)

        settings_button = self._button(NSMakeRect(660, 382, 62, 34), "设置", "openSettings:")
        content.addSubview_(settings_button)

        quit_button = self._button(NSMakeRect(732, 382, 60, 34), "退出", "quitApp:")
        content.addSubview_(quit_button)

        scroll = NSScrollView.alloc().initWithFrame_(NSMakeRect(28, 28, 764, 334))
        scroll.setHasVerticalScroller_(True)
        scroll.setHasHorizontalScroller_(False)
        self.log_view = NSTextView.alloc().initWithFrame_(NSMakeRect(0, 0, 764, 334))
        self.log_view.setEditable_(False)
        self.log_view.setString_("等待开始。\n")
        self.log_view.setTextColor_(NSColor.textColor())
        scroll.setDocumentView_(self.log_view)
        content.addSubview_(scroll)

        self.window.makeKeyAndOrderFront_(None)
        NSApplication.sharedApplication().activateIgnoringOtherApps_(True)

    @objc.python_method
    def _mode_tab(self, identifier: str, title: str, detail: str):
        item = NSTabViewItem.alloc().initWithIdentifier_(identifier)
        item.setLabel_(title)
        view = NSTextField.alloc().initWithFrame_(NSMakeRect(16, 16, 720, 30))
        view.setStringValue_(detail)
        view.setEditable_(False)
        view.setBordered_(False)
        view.setDrawsBackground_(False)
        item.setView_(view)
        return item

    @objc.python_method
    def _label(self, frame, text: str):
        label = NSTextField.alloc().initWithFrame_(frame)
        label.setStringValue_(text)
        label.setEditable_(False)
        label.setBordered_(False)
        label.setDrawsBackground_(False)
        return label

    @objc.python_method
    def _button(self, frame, title: str, action: str):
        button = NSButton.alloc().initWithFrame_(frame)
        button.setTitle_(title)
        button.setBezelStyle_(NSBezelStyleRounded)
        button.setTarget_(self)
        button.setAction_(action)
        return button

    def chooseFolder_(self, sender):
        panel = NSOpenPanel.openPanel()
        panel.setCanChooseFiles_(False)
        panel.setCanChooseDirectories_(True)
        panel.setAllowsMultipleSelection_(False)
        if panel.runModal():
            url = panel.URLs()[0]
            self.path_field.setStringValue_(str(url.path()))

    def startProcessing_(self, sender):
        if self.processing:
            return
        folder_input = str(self.path_field.stringValue()).strip()
        folder = normalize_folder_input(folder_input)
        if not folder.exists() or not folder.is_dir():
            self._show_alert("路径无效", "请选择一个有效的发票文件夹。")
            return

        self.processing = True
        self.excel_path = None
        self.start_button.setEnabled_(False)
        self.open_excel_button.setEnabled_(False)
        self.log_view.setString_("")

        mode = str(self.tab_view.selectedTabViewItem().identifier())
        worker = threading.Thread(target=self._process_folder, args=(folder, mode), daemon=True)
        worker.start()

    def openExcel_(self, sender):
        if self.excel_path:
            NSWorkspace.sharedWorkspace().openFile_(str(self.excel_path))

    def quitApp_(self, sender):
        NSApplication.sharedApplication().terminate_(None)

    def openSettings_(self, sender):
        self._show_settings_window()

    def saveSettings_(self, sender):
        self.settings.filename_format = str(self.format_field.stringValue()).strip() or AppSettings.filename_format
        self.settings.auto_update_rates = bool(self.auto_rates_button.state())
        self.settings.annotate_converted = bool(self.annotate_button.state())
        try:
            self.settings.sequence_start = max(1, int(str(self.seq_field.stringValue()).strip() or "1"))
        except ValueError:
            self.settings.sequence_start = 1
        self.settings.save()
        self._append_log_main("设置已保存。")

    def updateRates_(self, sender):
        self._append_log_main("正在更新汇率...")
        threading.Thread(target=self._manual_rate_update, daemon=True).start()

    @objc.python_method
    def _show_settings_window(self):
        if self.settings_window:
            self.settings_window.makeKeyAndOrderFront_(None)
            return

        style = NSWindowStyleMaskTitled | NSWindowStyleMaskClosable
        self.settings_window = NSWindow.alloc().initWithContentRect_styleMask_backing_defer_(
            NSMakeRect(0, 0, 520, 270),
            style,
            NSBackingStoreBuffered,
            False,
        )
        self.settings_window.setTitle_("设置")
        self.settings_window.center()
        content = self.settings_window.contentView()

        content.addSubview_(self._label(NSMakeRect(24, 220, 150, 22), "文件名格式"))
        self.format_field = NSTextField.alloc().initWithFrame_(NSMakeRect(120, 216, 376, 28))
        self.format_field.setStringValue_(self.settings.filename_format)
        content.addSubview_(self.format_field)

        help_text = self._label(NSMakeRect(120, 194, 376, 18), "可用：{yymm} {seq} {category} {currency} {amount} {merchant}")
        help_text.setTextColor_(NSColor.secondaryLabelColor())
        content.addSubview_(help_text)

        content.addSubview_(self._label(NSMakeRect(24, 154, 150, 22), "起始序号"))
        self.seq_field = NSTextField.alloc().initWithFrame_(NSMakeRect(120, 150, 90, 28))
        self.seq_field.setStringValue_(str(self.settings.sequence_start))
        content.addSubview_(self.seq_field)

        self.auto_rates_button = NSButton.alloc().initWithFrame_(NSMakeRect(120, 112, 260, 22))
        self.auto_rates_button.setButtonType_(NSButtonTypeSwitch)
        self.auto_rates_button.setTitle_("打开 App 时自动更新当天汇率")
        self.auto_rates_button.setState_(1 if self.settings.auto_update_rates else 0)
        content.addSubview_(self.auto_rates_button)

        self.annotate_button = NSButton.alloc().initWithFrame_(NSMakeRect(120, 82, 260, 22))
        self.annotate_button.setButtonType_(NSButtonTypeSwitch)
        self.annotate_button.setTitle_("外币换算后在票据右上角标注")
        self.annotate_button.setState_(1 if self.settings.annotate_converted else 0)
        content.addSubview_(self.annotate_button)

        update_button = self._button(NSMakeRect(120, 34, 116, 32), "更新汇率", "updateRates:")
        content.addSubview_(update_button)
        save_button = self._button(NSMakeRect(380, 34, 116, 32), "保存", "saveSettings:")
        content.addSubview_(save_button)

        self.settings_window.makeKeyAndOrderFront_(None)

    @objc.python_method
    def _process_folder(self, folder: Path, mode: str):
        logs: list[str] = []
        results = []
        report_currency = "HKD" if mode == "HK" else "CNY"
        try:
            if self.settings.auto_update_rates:
                _, message = self.rates.update_if_needed()
                self._append_log_main(message)
                AppHelper.callAfter(self.rate_label.setStringValue_, self.rates.short_status())

            files = iter_invoice_files(folder)
            if not files:
                self._append_log_main("未找到 PDF、JPG、PNG 发票文件。")
                return

            self._append_log_main(f"{mode} 模式：找到 {len(files)} 个待处理文件。")
            start_seq = max(1, self.settings.sequence_start)
            for idx, path in enumerate(files, 1):
                seq = start_seq + idx - 1
                self._append_log_main(f"[{idx}/{len(files)}] 正在处理：{path.name}")
                try:
                    result = process_one_file(path, logs, seq=seq, mode=mode, settings=self.settings, rates=self.rates)
                    if result:
                        results.append(result)
                        self._append_log_main(
                            f"  - 完成：{result.new_name}（{result.original_currency} {result.amount_original:g} -> {result.report_currency} {result.report_amount}）"
                        )
                except Exception as exc:
                    self._append_log_main(f"  - 跳过：{path.name}，原因：{exc}")
                    self._append_log_main(traceback.format_exc(limit=2).strip())

                while logs:
                    self._append_log_main(logs.pop(0))

            if not results:
                self._append_log_main("处理结束，但没有成功解析的发票。")
                return

            excel_path = build_excel(results, folder, report_currency=report_currency)
            self.excel_path = excel_path
            self._append_log_main(f"全部完成。已生成：{excel_path.name}")
            AppHelper.callAfter(self.open_excel_button.setEnabled_, True)
        finally:
            self.processing = False
            AppHelper.callAfter(self.start_button.setEnabled_, True)

    @objc.python_method
    def _refresh_rates_on_launch(self):
        _, message = self.rates.update_if_needed()
        self._append_log_main(message)
        AppHelper.callAfter(self.rate_label.setStringValue_, self.rates.short_status())

    @objc.python_method
    def _manual_rate_update(self):
        _, message = self.rates.update_if_needed(force=True)
        self._append_log_main(message)
        AppHelper.callAfter(self.rate_label.setStringValue_, self.rates.short_status())

    @objc.python_method
    def _append_log_main(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        AppHelper.callAfter(self._append_log, f"[{timestamp}] {message}")

    @objc.python_method
    def _append_log(self, message: str):
        current = str(self.log_view.string())
        self.log_view.setString_(current + message + "\n")
        self.log_view.scrollRangeToVisible_((len(self.log_view.string()), 0))

    @objc.python_method
    def _show_alert(self, title: str, message: str):
        alert = NSAlert.alloc().init()
        alert.setMessageText_(title)
        alert.setInformativeText_(message)
        alert.runModal()


def main():
    app = NSApplication.sharedApplication()
    app.setActivationPolicy_(NSApplicationActivationPolicyRegular)
    delegate = InvoiceNativeApp.alloc().init()
    app.setDelegate_(delegate)
    AppHelper.runEventLoop()


if __name__ == "__main__":
    main()
