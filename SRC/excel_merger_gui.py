"""
带颜色高亮的Excel文件合并工具（图形界面版）
功能：比较两个Excel文件的第一列（商品名称），将文件2中独有的行添加到文件1，合并重复项的订单量
新增功能：合并结果中新增项用黄色标记，合并项用绿色标记
依赖：pandas, openpyxl, tkinter（均为标准库或已安装库）
使用方法：双击运行此文件
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import threading
import urllib.request
import urllib.error
import urllib.parse
import json
import sys
import subprocess
import shutil
from copy import copy

# ===== 版本与更新配置 =====
# 当前版本号，需与远端 version.json 中的 version 字段格式一致
CURRENT_VERSION = "v2.26"

# ===== 更新源选择（宏） =====
# 可选值："github" 或 "gitee"
# 国内网络推荐改为 "gitee"，海外或能访问 GitHub 的环境使用 "github"
UPDATE_SOURCE = "gitee"

# ----- GitHub 配置 -----
GITHUB_REPO = "zemingzhang/537update"
# 通过 GitHub API 自动获取默认分支名（无需手动指定 main/master）
GITHUB_API_REPO_URL = f"https://api.github.com/repos/{GITHUB_REPO}"

# ----- Gitee 配置 -----
# 格式：你的 Gitee 用户名/仓库名
GITEE_REPO = "zemingzhang/537update"
# Gitee 原始文件 CDN，用于读取 version.json
GITEE_RAW_BASE = f"https://gitee.com/{GITEE_REPO}/raw/master"

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("烟草订单文件合并工具 v2.26 ")
        self.root.geometry("780x720")
        
        # 设置窗口图标（可选）
        try:
            self.root.iconbitmap(default='')
        except:
            pass
        
        # 初始化文件路径变量
        self.file1_path = ""
        self.file2_path = ""
        
        # 颜色定义
        self.new_item_color = PatternFill(start_color="FFFF00",  # 黄色
                                          end_color="FFFF00", 
                                          fill_type="solid")
        self.merged_item_color = PatternFill(start_color="90EE90",  # 浅绿色
                                             end_color="90EE90", 
                                             fill_type="solid")
        
        # 创建界面组件
        self.setup_ui()
    
    def setup_ui(self):
        """创建和布局所有界面组件"""
        # ===== 顶部更新栏 =====
        top_bar = tk.Frame(self.root, bg="#e8e8e8", bd=1, relief=tk.GROOVE)
        top_bar.pack(fill=tk.X, padx=0, pady=(0, 5))

        self.update_version_label = tk.Label(
            top_bar,
            text=f"当前版本: {CURRENT_VERSION}",
            font=("微软雅黑", 8),
            fg="gray",
            bg="#e8e8e8"
        )
        self.update_version_label.pack(side=tk.LEFT, padx=10, pady=4)

        self.update_btn = ttk.Button(
            top_bar,
            text="🔄 检查更新",
            command=self.check_for_updates,
            width=12
        )
        self.update_btn.pack(side=tk.RIGHT, padx=10, pady=3)

        # ===== 页面导航栏 =====
        self.nav_bar = tk.Frame(self.root, bg="#f0f0f0", bd=0)
        self.nav_bar.pack(fill=tk.X, padx=0, pady=0)

        self.tab_btns = []
        tab_defs = [("📋  订单合并", 0), ("📊  档位投放比较", 1), ("�  比较订货", 2), ("🛠  其它功能", 3)]
        for name, idx in tab_defs:
            btn = tk.Button(
                self.nav_bar,
                text=name,
                font=("微软雅黑", 10),
                relief=tk.FLAT,
                bd=0,
                padx=18,
                pady=7,
                cursor="hand2",
                command=lambda i=idx: self.switch_page(i)
            )
            btn.pack(side=tk.LEFT)
            self.tab_btns.append(btn)

        # 导航栏底部分隔线
        tk.Frame(self.root, bg="#cccccc", height=1).pack(fill=tk.X)

        # ===== 底部状态栏（先 pack 确保布局正确）=====
        self.status_bar = tk.Label(self.root,
                                   text="等待操作...",
                                   font=("微软雅黑", 9),
                                   fg="gray",
                                   bd=1,
                                   relief=tk.SUNKEN,
                                   anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # ===== 内容容器 =====
        self.content_container = tk.Frame(self.root)
        self.content_container.pack(fill=tk.BOTH, expand=True)

        # 页面1：Excel合并
        self.page1 = tk.Frame(self.content_container)
        self.page1.pack(fill=tk.BOTH, expand=True)
        self._setup_page1()

        # 页面2：档位投放比较
        self.page2 = tk.Frame(self.content_container)
        self._setup_page2()

        # 页面3：比较订货
        self.page3 = tk.Frame(self.content_container)
        self._setup_page3()

        # 页面4：其它功能
        self.page4 = tk.Frame(self.content_container)
        self._setup_page4()

        # 默认显示第一页
        self.current_page = 0
        self._update_tab_styles()

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _setup_page1(self):
        """创建 Excel 合并页面的内容"""
        # 标题区域
        title_frame = tk.Frame(self.page1)
        title_frame.pack(pady=(15, 5), fill=tk.X)

        tk.Label(title_frame, text="🎨 订单文件合并工具",
                 font=("微软雅黑", 18, "bold")).pack()
        tk.Label(title_frame,
                 text="请选择需要合并的Excel文件（需包含'商品'、'盒码'、'订单量'三列）",
                 font=("微软雅黑", 10), wraplength=500, fg="gray").pack(pady=5)

        # 模式选择栏
        mode_frame = tk.Frame(self.page1, bg="#f0f0f0", relief=tk.GROOVE, bd=1)
        mode_frame.pack(fill=tk.X, padx=20, pady=(0, 5))
        tk.Label(mode_frame, text="合并模式：",
                 font=("微软雅黑", 10), bg="#f0f0f0").pack(
                     side=tk.LEFT, padx=10, pady=5)
        self.p1_mode = tk.StringVar(value="two")
        ttk.Radiobutton(mode_frame, text="双文件比较（带颜色标注）",
                        variable=self.p1_mode, value="two",
                        command=self._p1_switch_mode).pack(
                            side=tk.LEFT, padx=8, pady=5)
        ttk.Radiobutton(mode_frame, text="多文件合并（不标颜色）",
                        variable=self.p1_mode, value="multi",
                        command=self._p1_switch_mode).pack(
                            side=tk.LEFT, padx=8, pady=5)

        # 文件选择容器（双文件 / 多文件区域在此切换）
        self.p1_filesel_container = tk.Frame(self.page1)
        self.p1_filesel_container.pack(fill=tk.X)

        # ── 双文件区域（默认显示）──────────────────────────────────
        self.p1_twofile_frame = tk.Frame(self.p1_filesel_container)
        self.p1_twofile_frame.pack(fill=tk.X)

        color_frame = tk.Frame(self.p1_twofile_frame)
        color_frame.pack(pady=5, padx=20, fill=tk.X)
        tk.Label(color_frame, text="颜色说明：",
                 font=("微软雅黑", 10, "bold")).pack(anchor=tk.W)

        new_cf = tk.Frame(color_frame)
        new_cf.pack(anchor=tk.W, pady=2)
        tk.Label(new_cf, text="  ", bg="yellow", width=3, height=1).pack(
            side=tk.LEFT, padx=(0, 5))
        tk.Label(new_cf, text="新增项（文件2中有，文件1中没有的商品）",
                 font=("微软雅黑", 9)).pack(side=tk.LEFT)

        merged_cf = tk.Frame(color_frame)
        merged_cf.pack(anchor=tk.W, pady=2)
        tk.Label(merged_cf, text="  ", bg="#90EE90", width=3, height=1).pack(
            side=tk.LEFT, padx=(0, 5))
        tk.Label(merged_cf, text="合并项（两个文件中都存在的商品，订单量已相加）",
                 font=("微软雅黑", 9)).pack(side=tk.LEFT)

        self.create_file_selection("第一个Excel文件", 0, self.p1_twofile_frame)
        self.create_file_selection("第二个Excel文件", 1, self.p1_twofile_frame)

        # ── 多文件区域（初始隐藏）──────────────────────────────────
        self.p1_multifile_frame = tk.Frame(self.p1_filesel_container)

        files_lf = tk.LabelFrame(
            self.p1_multifile_frame,
            text="选择文件（数量不限，以第一个文件为基础依次合并）",
            font=("微软雅黑", 10))
        files_lf.pack(padx=20, pady=6, fill=tk.X)

        self.p1_multi_file_paths = []

        btn_row = tk.Frame(files_lf)
        btn_row.pack(fill=tk.X, padx=8, pady=(6, 2))
        ttk.Button(btn_row, text="➕ 添加文件…",
                   command=self._p1_add_files,
                   width=14).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="➖ 删除选中",
                   command=self._p1_remove_selected,
                   width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="清除全部",
                   command=self._p1_clear_multifiles,
                   width=10).pack(side=tk.LEFT)
        self.p1_multi_count_label = tk.Label(
            btn_row, text="已选 0 个文件",
            font=("微软雅黑", 9), fg="gray")
        self.p1_multi_count_label.pack(side=tk.RIGHT, padx=8)

        list_frame = tk.Frame(files_lf)
        list_frame.pack(fill=tk.X, padx=8, pady=(2, 8))
        list_sb = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        list_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.p1_multi_listbox = tk.Listbox(
            list_frame, height=4, font=("微软雅黑", 9),
            selectmode=tk.EXTENDED, yscrollcommand=list_sb.set,
            activestyle="dotbox", bg="#fafafa")
        self.p1_multi_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        list_sb.config(command=self.p1_multi_listbox.yview)

        # ── 状态显示区域 ────────────────────────────────────────────
        status_frame = tk.LabelFrame(self.page1, text="处理状态",
                                      font=("微软雅黑", 10))
        status_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(status_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.status_text = tk.Text(status_frame, height=8, width=70,
                                   font=("宋体", 9),
                                   yscrollcommand=scrollbar.set)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH,
                              expand=True, padx=5, pady=5)
        scrollbar.config(command=self.status_text.yview)

        self.update_status("就绪。请选择Excel文件。\n")
        self.status_text.config(state=tk.DISABLED)

        # 按钮区域
        button_frame = tk.Frame(self.page1)
        button_frame.pack(pady=10)

        style = ttk.Style()
        style.configure("Large.TButton", font=("微软雅黑", 12), padding=10)

        self.merge_button = ttk.Button(button_frame,
                                       text="开始合并文件",
                                       command=self.merge_files,
                                       state=tk.DISABLED)
        self.merge_button.pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="退出程序",
                   command=self.root.quit).pack(side=tk.LEFT, padx=5)

    # ===== Page1 多文件合并辅助方法 =====

    def _p1_switch_mode(self):
        """切换双文件/多文件模式"""
        if self.p1_mode.get() == "two":
            self.p1_multifile_frame.pack_forget()
            self.p1_twofile_frame.pack(fill=tk.X)
            self.check_files_selected()
        else:
            self.p1_twofile_frame.pack_forget()
            self.p1_multifile_frame.pack(fill=tk.X)
            self._p1_check_multi_ready()

    def _p1_add_files(self):
        """多文件模式：添加文件（自动去重）"""
        paths = filedialog.askopenfilenames(
            title="选择文件（可多选）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        added = 0
        for path in paths:
            if path not in self.p1_multi_file_paths:
                self.p1_multi_file_paths.append(path)
                self.p1_multi_listbox.insert(tk.END, os.path.basename(path))
                added += 1
        if added:
            self.update_status(
                f"已添加 {added} 个文件，当前共 {len(self.p1_multi_file_paths)} 个")
        self._p1_update_multi_count()

    def _p1_remove_selected(self):
        """多文件模式：删除选中文件"""
        selected = sorted(
            self.p1_multi_listbox.curselection(), reverse=True)
        for i in selected:
            self.p1_multi_file_paths.pop(i)
            self.p1_multi_listbox.delete(i)
        if selected:
            self.update_status(
                f"已删除 {len(selected)} 个，剩余 {len(self.p1_multi_file_paths)} 个")
        self._p1_update_multi_count()

    def _p1_clear_multifiles(self):
        """多文件模式：清除全部"""
        self.p1_multi_file_paths.clear()
        self.p1_multi_listbox.delete(0, tk.END)
        self._p1_update_multi_count()
        self.update_status("已清除所有文件选择。")

    def _p1_update_multi_count(self):
        """更新多文件计数标签并刷新按钮状态"""
        n = len(self.p1_multi_file_paths)
        self.p1_multi_count_label.config(
            text=f"已选 {n} 个文件",
            fg="#0066cc" if n > 0 else "gray"
        )
        self._p1_check_multi_ready()

    def _p1_check_multi_ready(self):
        """多文件模式：>=2 个文件时才能合并"""
        n = len(self.p1_multi_file_paths)
        if n >= 2:
            self.merge_button.config(state=tk.NORMAL)
            self.status_bar.config(text=f"已选 {n} 个文件，可以开始合并")
        else:
            self.merge_button.config(state=tk.DISABLED)
            self.status_bar.config(text="多文件模式：请至少选择 2 个文件")

    def _p1_merge_multi(self):
        """多文件合并核心逻辑（不标颜色）"""
        files = list(self.p1_multi_file_paths)
        if len(files) < 2:
            messagebox.showerror("错误", "请至少选择 2 个文件！")
            return

        self.merge_button.config(state=tk.DISABLED, text="处理中...")
        self.status_bar.config(text="正在处理，请稍候...")
        self.root.update()

        try:
            required_columns = ['商品', '盒码', '订单量']
            self.update_status(f"多文件模式：共 {len(files)} 个文件")

            dfs = []
            for path in files:
                df = pd.read_excel(path)
                missing = [c for c in required_columns if c not in df.columns]
                if missing:
                    raise ValueError(
                        f"文件「{os.path.basename(path)}」缺少必要列: {missing}")
                dfs.append(df)
                self.update_status(
                    f"  已读取: {os.path.basename(path)}，共 {len(df)} 行")

            result_df = dfs[0].copy()

            for i, df_next in enumerate(dfs[1:], start=2):
                result_prods = set(result_df['商品'].astype(str).str.strip())
                next_prods   = set(df_next['商品'].astype(str).str.strip())

                # 新增商品
                new_prods = next_prods - result_prods
                if new_prods:
                    new_rows = df_next[
                        df_next['商品'].astype(str).str.strip().isin(new_prods)
                    ].copy()
                    result_df = pd.concat(
                        [result_df, new_rows], ignore_index=True, sort=False)
                    self.update_status(
                        f"  文件{i}: 新增 {len(new_prods)} 个商品")

                # 共同商品合并订单量
                common_prods = result_prods.intersection(next_prods)
                qty_map = {}
                for _, row in df_next.iterrows():
                    pname = str(row['商品']).strip()
                    if pname in common_prods:
                        try:
                            qty_map[pname] = qty_map.get(pname, 0) + float(row['订单量'])
                        except (ValueError, TypeError):
                            pass

                merge_count = 0
                for idx, row in result_df.iterrows():
                    pname = str(row['商品']).strip()
                    if pname in qty_map:
                        result_df.at[idx, '订单量'] = (
                            float(row['订单量']) + qty_map[pname])
                        merge_count += 1

                if merge_count:
                    self.update_status(
                        f"  文件{i}: 合并 {merge_count} 个共同商品的订单量")

            self.update_status(f"✅ 合并完成，共 {len(result_df)} 行")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"多文件合并结果_{timestamp}.xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            if not out_path:
                self.update_status("操作已取消")
                self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
                self.status_bar.config(text="操作取消")
                return

            result_df[['商品', '盒码', '订单量']].to_excel(out_path, index=False)
            self.update_status(f"   已保存至: {out_path}")
            self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
            self.status_bar.config(
                text=f"完成！保存为: {os.path.basename(out_path)}")

            if messagebox.askyesno(
                "完成",
                f"多文件合并完成！\n共 {len(result_df)} 行\n是否打开结果文件？"
            ):
                try:
                    os.startfile(out_path)
                except Exception:
                    pass

        except Exception as e:
            self.update_status(f"❌ 错误: {str(e)}")
            messagebox.showerror("处理错误", str(e))
            self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
            self.status_bar.config(text="处理出错，请检查文件格式")

    def _setup_page2(self):
        """创建档位投放比较页面的内容"""
        # 标题
        title_frame = tk.Frame(self.page2)
        title_frame.pack(pady=(15, 5), fill=tk.X)

        tk.Label(title_frame, text="📊 档位投放比较",
                 font=("微软雅黑", 16, "bold")).pack()
        tk.Label(title_frame,
                 text='读取Excel文件，按选定档位保留"是"的行，删除所有档位列，输出结果',
                 font=("微软雅黑", 9), fg="gray").pack(pady=3)

        # 文件选择区（数量不限）
        files_lf = tk.LabelFrame(self.page2, text="选择文件（数量不限，支持多选）",
                                 font=("微软雅黑", 10))
        files_lf.pack(padx=20, pady=6, fill=tk.X)

        self.p2_file_paths = []

        # 操作按钮行
        btn_row = tk.Frame(files_lf)
        btn_row.pack(fill=tk.X, padx=8, pady=(6, 2))

        ttk.Button(btn_row, text="➕ 添加文件…",
                   command=self._p2_add_files,
                   width=14).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="➖ 删除选中",
                   command=self._p2_remove_selected,
                   width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="清除全部",
                   command=self._p2_clear_files,
                   width=10).pack(side=tk.LEFT)

        self.p2_file_count_label = tk.Label(btn_row, text="已选 0 个文件",
                                             font=("微软雅黑", 9), fg="gray")
        self.p2_file_count_label.pack(side=tk.RIGHT, padx=8)

        # 文件列表框
        list_frame = tk.Frame(files_lf)
        list_frame.pack(fill=tk.X, padx=8, pady=(2, 8))

        list_sb = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
        list_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.p2_file_listbox = tk.Listbox(
            list_frame, height=4,
            font=("微软雅黑", 9),
            selectmode=tk.EXTENDED,
            yscrollcommand=list_sb.set,
            activestyle="dotbox",
            bg="#fafafa"
        )
        self.p2_file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        list_sb.config(command=self.p2_file_listbox.yview)

        # 档位选择 + 处理按钮
        opt_frame = tk.Frame(self.page2)
        opt_frame.pack(padx=20, pady=8, fill=tk.X)

        tk.Label(opt_frame, text="选择档位：",
                 font=("微软雅黑", 10)).pack(side=tk.LEFT)
        self.p2_gear_var = tk.StringVar(value="1")
        ttk.Combobox(opt_frame, textvariable=self.p2_gear_var,
                     values=[str(i) for i in range(1, 31)],
                     width=6, state="readonly").pack(side=tk.LEFT, padx=(0, 20))

        self.p2_process_btn = ttk.Button(opt_frame, text="开始处理",
                                         command=self._p2_process)
        self.p2_process_btn.pack(side=tk.LEFT, padx=5)

        # 状态日志区
        status_lf = tk.LabelFrame(self.page2, text="处理状态",
                                  font=("微软雅黑", 10))
        status_lf.pack(padx=20, pady=6, fill=tk.BOTH, expand=True)

        sb = tk.Scrollbar(status_lf)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.p2_status = tk.Text(status_lf, height=8,
                                 font=("宋体", 9),
                                 yscrollcommand=sb.set)
        self.p2_status.pack(side=tk.LEFT, fill=tk.BOTH,
                            expand=True, padx=5, pady=5)
        sb.config(command=self.p2_status.yview)
        self._p2_log("就绪。请选择文件并选择档位。\n")
        self.p2_status.config(state=tk.DISABLED)

    # ===== 档位投放比较页面辅助方法 =====

    # 阿拉伯数字 → 中文档位列名映射（文件中第3行的列标题格式）
    _P2_GEAR_CN = {
        1: '一档', 2: '二档', 3: '三档', 4: '四档', 5: '五档',
        6: '六档', 7: '七档', 8: '八档', 9: '九档', 10: '十档',
        11: '十一档', 12: '十二档', 13: '十三档', 14: '十四档', 15: '十五档',
        16: '十六档', 17: '十七档', 18: '十八档', 19: '十九档', 20: '二十档',
        21: '二十一档', 22: '二十二档', 23: '二十三档', 24: '二十四档', 25: '二十五档',
        26: '二十六档', 27: '二十七档', 28: '二十八档', 29: '二十九档', 30: '三十档',
    }

    def _p2_add_files(self):
        """打开多文件对话框，将选中文件追加到列表（自动去重）"""
        paths = filedialog.askopenfilenames(
            title="选择文件（可多选）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        added = 0
        for path in paths:
            if path not in self.p2_file_paths:
                self.p2_file_paths.append(path)
                self.p2_file_listbox.insert(tk.END, os.path.basename(path))
                added += 1
        if added:
            self._p2_log(
                f"已添加 {added} 个文件，当前共 {len(self.p2_file_paths)} 个")
        self._p2_update_file_count()

    def _p2_remove_selected(self):
        """删除列表中高亮选中的文件（支持多选）"""
        selected = sorted(self.p2_file_listbox.curselection(), reverse=True)
        for i in selected:
            self.p2_file_paths.pop(i)
            self.p2_file_listbox.delete(i)
        if selected:
            self._p2_log(
                f"已删除 {len(selected)} 个文件，剩余 {len(self.p2_file_paths)} 个")
        self._p2_update_file_count()

    def _p2_clear_files(self):
        """清除所有文件"""
        self.p2_file_paths.clear()
        self.p2_file_listbox.delete(0, tk.END)
        self._p2_update_file_count()
        self._p2_log("已清除所有文件选择。")

    def _p2_update_file_count(self):
        """更新文件计数标签颜色"""
        n = len(self.p2_file_paths)
        self.p2_file_count_label.config(
            text=f"已选 {n} 个文件",
            fg="#0066cc" if n > 0 else "gray"
        )

    def _p2_log(self, msg):
        """向状态框追加日志"""
        self.p2_status.config(state=tk.NORMAL)
        self.p2_status.insert(tk.END, f"> {msg}\n")
        self.p2_status.see(tk.END)
        self.p2_status.config(state=tk.DISABLED)
        self.root.update()

    def _p2_gear_indices(self, raw):
        """
        从原始 DataFrame 第3行（index=2）识别所有档位列。
        返回 dict: 档位名称 -> 列索引(int)
        文件结构：行0=大标题, 行1=主列头, 行2=档位子列头(三十档…一档), 行3+=数据
        """
        header2 = raw.iloc[2].tolist()
        result = {}
        for i, v in enumerate(header2):
            v_str = str(v).strip()
            if '档' in v_str and v_str not in ('nan', 'NaN', ''):
                result[v_str] = i
        return result

    def _p2_apply_filter(self, raw, all_gear_idxs, sel_idx, keep_products=None):
        """
        对原始 DataFrame 执行列删除和行筛选，前3行表头结构原样保留。

        列处理：删除 sel_idx 以外的所有档位列（三行表头也同步删除）
        行处理（数据行 index>=3）：
          keep_products=None  → 单文件模式：保留选中档位列值="是"的行
          keep_products=set   → 多文件模式：保留商品名称（列1）在集合中的行
        """
        drop_idxs = sorted([i for i in all_gear_idxs if i != sel_idx])

        # 删除其他档位列（表头与数据行同步）
        raw_trimmed = raw.drop(columns=drop_idxs)
        raw_trimmed = raw_trimmed.reset_index(drop=True)
        raw_trimmed.columns = range(raw_trimmed.shape[1])

        # 删除列后，选中档位列始终位于位置4（档位列原先连续占据列4-33）
        gear_pos = 4

        header_rows = raw_trimmed.iloc[:3].copy()
        data_rows = raw_trimmed.iloc[3:].copy()

        # 去除商品名称为空的无效行
        valid_mask = ~data_rows.iloc[:, 1].astype(str).str.strip().isin(
            ['nan', 'NaN', ''])
        data_rows = data_rows[valid_mask]

        if keep_products is None:
            # 单文件：选中档位列="是"的行保留，"否"及其他删除
            mask = data_rows.iloc[:, gear_pos].astype(str).str.strip() == '是'
        else:
            # 多文件：商品名称在跨文件"是"集合中则保留
            mask = data_rows.iloc[:, 1].astype(str).str.strip().isin(
                keep_products)

        before = len(data_rows)
        data_filtered = data_rows[mask].copy()
        self._p2_log(f"  保留 {len(data_filtered)}/{before} 条，"
                     f"删除 {before - len(data_filtered)} 条")

        # 多文件模式：保留行中选中档位列仍为"否"的改为"是"
        if keep_products is not None:
            no_mask = data_filtered.iloc[:, gear_pos].astype(str).str.strip() == '否'
            changed_count = int(no_mask.sum())
            if changed_count > 0:
                data_filtered.iloc[no_mask.values, gear_pos] = '是'
                self._p2_log(f"  已将 {changed_count} 条「否」→「是」"
                             f"（该商品在其他文件中有投放）")

        return pd.concat([header_rows, data_filtered], ignore_index=True)

    def _p2_process(self):
        """执行档位筛选处理"""
        selected_files = list(self.p2_file_paths)
        if not selected_files:
            messagebox.showerror("错误", "请至少选择一个Excel文件！",
                                 parent=self.root)
            return

        gear_num = int(self.p2_gear_var.get())
        gear_name = self._P2_GEAR_CN[gear_num]
        self.p2_process_btn.config(state=tk.DISABLED, text="处理中...")

        try:
            file_raws = []   # (path, raw_df, gear_map, sel_idx)

            # ── 读取全部文件，校验档位列存在 ──────────────────────────
            for path in selected_files:
                self._p2_log(f"正在读取: {os.path.basename(path)}")
                raw = pd.read_excel(path, header=None, dtype=str)
                self._p2_log(
                    f"  共 {raw.shape[0]} 行 × {raw.shape[1]} 列")

                gear_map = self._p2_gear_indices(raw)
                if gear_name not in gear_map:
                    raise ValueError(
                        f"文件「{os.path.basename(path)}」中未找到"
                        f"「{gear_name}」列\n"
                        f"检测到的档位列：{list(gear_map.keys())}"
                    )

                sel_idx = gear_map[gear_name]
                self._p2_log(f"  「{gear_name}」→ 第{sel_idx + 1}列")
                file_raws.append((path, raw, gear_map, sel_idx))

            # ── 按模式过滤 ────────────────────────────────────────────
            if len(file_raws) == 1:
                # 单文件：删除其余档位列，仅保留选中档位="是"的行
                path, raw, gear_map, sel_idx = file_raws[0]
                all_gear_idxs = list(gear_map.values())
                self._p2_log(
                    f"单文件模式：删除其余 {len(all_gear_idxs) - 1} 个档位列，"
                    f'保留「{gear_name}」="是"的行')
                result = self._p2_apply_filter(
                    raw, all_gear_idxs, sel_idx, keep_products=None)
            else:
                # ── 多文件模式 ────────────────────────────────────────
                # 第一个文件作为基础：单文件过滤（只保留 gear="是" 的行）
                path0, raw0, gear_map0, sel_idx0 = file_raws[0]
                all_gear_idxs0 = list(gear_map0.values())
                self._p2_log(
                    f"多文件模式：以「{os.path.basename(path0)}」为基础（档位列保留「是」的行）")
                result = self._p2_apply_filter(
                    raw0, all_gear_idxs0, sel_idx0, keep_products=None)

                # 收集第一个文件中 所有 商品名称（含"否"行）作为已知集合
                accumulated = set(
                    raw0.iloc[3:, 1].astype(str).str.strip()
                ) - {'nan', 'NaN', ''}

                # 单独记录第一个文件中 gear="否" 的商品集合
                # 这些商品若在后续文件中出现 gear="是"，也需追加到输出
                _d0 = raw0.iloc[3:].copy()
                _d0 = _d0[~_d0.iloc[:, 1].astype(str).str.strip().isin(
                    ['nan', 'NaN', ''])]
                gear_yes_in_base = set(
                    _d0[_d0.iloc[:, sel_idx0].astype(str).str.strip() == '是']
                    .iloc[:, 1].astype(str).str.strip()
                )
                # gear_no_set：在前序文件中存在但 gear≠"是" 的商品
                gear_no_set = accumulated - gear_yes_in_base

                self._p2_log(
                    f'  基础文件共 {len(accumulated)} 个商品，'
                    f'其中 gear="是" {len(gear_yes_in_base)} 个，'
                    f'gear="否" {len(gear_no_set)} 个')

                # 后续文件：两种情况均追加
                #   ① 商品不在累积集合中 且 gear="是"（全新商品）
                #   ② 商品在前序文件中 gear="否"，当前文件 gear="是"（由否转是）
                gear_pos = 4
                for path, raw, gear_map, sel_idx in file_raws[1:]:
                    all_gear_idxs = list(gear_map.values())
                    drop_idxs = sorted(
                        [i for i in all_gear_idxs if i != sel_idx])

                    raw_trimmed = raw.drop(columns=drop_idxs)
                    raw_trimmed = raw_trimmed.reset_index(drop=True)
                    raw_trimmed.columns = range(raw_trimmed.shape[1])

                    data = raw_trimmed.iloc[3:].copy()
                    data = data[
                        ~data.iloc[:, 1].astype(str).str.strip().isin(
                            ['nan', 'NaN', ''])]

                    prod_col = data.iloc[:, 1].astype(str).str.strip()
                    gear_col = data.iloc[:, gear_pos].astype(str).str.strip()

                    mask_new      = ~prod_col.isin(accumulated) & (gear_col == '是')
                    mask_promoted = prod_col.isin(gear_no_set)  & (gear_col == '是')
                    new_rows = data[mask_new | mask_promoted].copy()

                    if len(new_rows) > 0:
                        n_new      = int(mask_new[mask_new | mask_promoted].sum())
                        n_promoted = int(mask_promoted[mask_new | mask_promoted].sum())
                        parts = []
                        if n_new:
                            parts.append(f"{n_new} 条全新商品")
                        if n_promoted:
                            parts.append(f'{n_promoted} 条（前序为"否"转"是"）')
                        self._p2_log(
                            f"  「{os.path.basename(path)}」追加 "
                            + "、".join(parts))

                        new_prods = set(
                            new_rows.iloc[:, 1].astype(str).str.strip())
                        accumulated.update(new_prods)
                        gear_no_set -= new_prods   # 已转"是"，移出候选集
                        result = pd.concat(
                            [result, new_rows], ignore_index=True)
                    else:
                        self._p2_log(
                            f"  「{os.path.basename(path)}」无新增商品")

            # ── 删除含"大连烟草"字样的首行（标题行）────────────────────
            first_row_text = " ".join(
                str(v) for v in result.iloc[0].tolist() if str(v) not in ('nan', 'NaN', ''))
            if "大连" in first_row_text or "烟草" in first_row_text:
                result = result.iloc[1:].reset_index(drop=True)
                self._p2_log('已删除首行标题（含"大连烟草"字样）')

            data_count = len(result) - 2  # 去掉2行表头（原3行减去已删的标题行）

            # ── 保存 ─────────────────────────────────────────────────
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"{gear_name}_投放结果_{timestamp}.xlsx"
            out_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            if not out_path:
                self._p2_log("操作已取消。")
                return

            result.to_excel(out_path, index=False, header=False)
            self._p2_log(f"✅ 已保存至: {out_path}")
            self._p2_log(
                f"   数据行: {data_count}，列数: {result.shape[1]}（其余档位列已删除）")

            if messagebox.askyesno(
                "完成",
                f"处理完成！\n档位：{gear_name}，共输出 {data_count} 条数据。\n"
                f"是否打开结果文件？",
                parent=self.root
            ):
                try:
                    os.startfile(out_path)
                except Exception:
                    pass

        except ValueError as e:
            self._p2_log(f"❌ 错误: {e}")
            messagebox.showerror("处理错误", str(e), parent=self.root)
        except Exception as e:
            self._p2_log(f"❌ 未知错误: {e}")
            messagebox.showerror("处理错误", str(e), parent=self.root)
        finally:
            self.p2_process_btn.config(state=tk.NORMAL, text="开始处理")

    def _setup_page3(self):
        """创建比较订货页面的内容"""
        title_frame = tk.Frame(self.page3)
        title_frame.pack(pady=20, fill=tk.X)

        tk.Label(title_frame, text="📋 比较订货",
                 font=("微软雅黑", 18, "bold")).pack()
        tk.Label(title_frame,
                 text="以文件1的商品名称为基准，与文件2的订货数据对比，标注匹配颜色",
                 font=("微软雅黑", 10), wraplength=500, fg="gray").pack(pady=5)

        # 颜色说明
        color_frame = tk.Frame(self.page3)
        color_frame.pack(pady=5, padx=20, fill=tk.X)
        tk.Label(color_frame, text="颜色说明：",
                 font=("微软雅黑", 10, "bold")).pack(anchor=tk.W)

        green_row = tk.Frame(color_frame)
        green_row.pack(anchor=tk.W, pady=2)
        tk.Label(green_row, text="  ", bg="#90EE90", width=3, height=1).pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(green_row, text="文件2中存在且订单量 > 0（匹配）",
                 font=("微软雅黑", 9)).pack(side=tk.LEFT)

        red_row = tk.Frame(color_frame)
        red_row.pack(anchor=tk.W, pady=2)
        tk.Label(red_row, text="  ", bg="#FF6B6B", width=3, height=1).pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(red_row, text="文件2中不存在的商品（未匹配）",
                 font=("微软雅黑", 9)).pack(side=tk.LEFT)

        none_row = tk.Frame(color_frame)
        none_row.pack(anchor=tk.W, pady=2)
        tk.Label(none_row, text="  ", bg="#f0f0f0", width=3, height=1,
                 relief=tk.GROOVE).pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(none_row, text="文件2中订单量 = 0 的商品（忽略，不标色）",
                 font=("微软雅黑", 9)).pack(side=tk.LEFT)

        # 文件1选择
        self.p3_file1_path = ""
        f1_lf = tk.LabelFrame(self.page3, text="选择文件1（投放结果文件）",
                               font=("微软雅黑", 10))
        f1_lf.pack(pady=8, padx=20, fill=tk.X)
        f1_inner = tk.Frame(f1_lf)
        f1_inner.pack(pady=8, padx=10, fill=tk.X)
        ttk.Button(f1_inner, text="浏览...", command=self._p3_select_file1,
                   width=15).pack(side=tk.LEFT, padx=(0, 10))
        self.p3_label1 = tk.Label(f1_inner, text="未选择文件",
                                   font=("微软雅黑", 9), fg="gray",
                                   bg="#f0f0f0", relief=tk.SUNKEN,
                                   anchor=tk.W, width=50)
        self.p3_label1.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 文件2选择
        self.p3_file2_path = ""
        f2_lf = tk.LabelFrame(self.page3, text="选择文件2（订货数据文件）",
                               font=("微软雅黑", 10))
        f2_lf.pack(pady=8, padx=20, fill=tk.X)
        f2_inner = tk.Frame(f2_lf)
        f2_inner.pack(pady=8, padx=10, fill=tk.X)
        ttk.Button(f2_inner, text="浏览...", command=self._p3_select_file2,
                   width=15).pack(side=tk.LEFT, padx=(0, 10))
        self.p3_label2 = tk.Label(f2_inner, text="未选择文件",
                                   font=("微软雅黑", 9), fg="gray",
                                   bg="#f0f0f0", relief=tk.SUNKEN,
                                   anchor=tk.W, width=50)
        self.p3_label2.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 状态区域
        status_frame = tk.LabelFrame(self.page3, text="处理状态",
                                      font=("微软雅黑", 10))
        status_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(status_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.p3_status_text = tk.Text(status_frame, height=8, width=70,
                                       font=("宋体", 9),
                                       yscrollcommand=scrollbar.set)
        self.p3_status_text.pack(side=tk.LEFT, fill=tk.BOTH,
                                  expand=True, padx=5, pady=5)
        scrollbar.config(command=self.p3_status_text.yview)
        self._p3_log("就绪。请选择两个文件。\n")
        self.p3_status_text.config(state=tk.DISABLED)

        # 按钮
        btn_frame = tk.Frame(self.page3)
        btn_frame.pack(pady=10)
        self.p3_compare_btn = ttk.Button(btn_frame, text="开始比较",
                                          command=self._p3_process,
                                          state=tk.DISABLED)
        self.p3_compare_btn.pack(side=tk.LEFT, padx=5)

    # ===== 比较订货页面辅助方法 =====

    def _p3_select_file1(self):
        path = filedialog.askopenfilename(
            title="选择文件1（投放结果文件）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.p3_file1_path = path
            self.p3_label1.config(text=os.path.basename(path), fg="black")
            self._p3_log(f"已选择文件1: {path}")
            self._p3_check_ready()

    def _p3_select_file2(self):
        path = filedialog.askopenfilename(
            title="选择文件2（订货数据文件）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if path:
            self.p3_file2_path = path
            self.p3_label2.config(text=os.path.basename(path), fg="black")
            self._p3_log(f"已选择文件2: {path}")
            self._p3_check_ready()

    def _p3_check_ready(self):
        if self.p3_file1_path and self.p3_file2_path:
            self.p3_compare_btn.config(state=tk.NORMAL)

    def _p3_log(self, msg):
        self.p3_status_text.config(state=tk.NORMAL)
        self.p3_status_text.insert(tk.END, f"> {msg}\n")
        self.p3_status_text.see(tk.END)
        self.p3_status_text.config(state=tk.DISABLED)
        self.root.update()

    def _p3_process(self):
        """比较订货核心处理"""
        self.p3_compare_btn.config(state=tk.DISABLED, text="处理中...")
        try:
            # ── 读取文件2，建立商品→订单量映射 ──────────────────────
            self._p3_log(f"正在读取文件2: {os.path.basename(self.p3_file2_path)}")
            df2 = pd.read_excel(self.p3_file2_path)

            prod_col2 = next(
                (c for c in df2.columns if '商品' in str(c)), None)
            qty_col2  = next(
                (c for c in df2.columns
                 if '订单' in str(c) or '数量' in str(c)), None)

            if prod_col2 is None:
                raise ValueError("文件2中未找到包含'商品'的列")
            if qty_col2 is None:
                raise ValueError("文件2中未找到订单/数量列")

            self._p3_log(f"  商品列:「{prod_col2}」 订单量列:「{qty_col2}」")

            order_map = {}
            for _, row in df2.iterrows():
                name = str(row[prod_col2]).strip()
                try:
                    qty = float(row[qty_col2])
                except (ValueError, TypeError):
                    qty = 0.0
                if name not in ('nan', 'NaN', ''):
                    order_map[name] = order_map.get(name, 0.0) + qty

            green_set = {k for k, v in order_map.items() if v > 0}
            zero_set  = {k for k, v in order_map.items() if v == 0}
            self._p3_log(
                f"  文件2：订单量>0 共 {len(green_set)} 个，"
                f"订单量=0（忽略）共 {len(zero_set)} 个")

            # ── 读取文件1，找商品名称列及数据起始行 ──────────────────
            self._p3_log(
                f"正在读取文件1: {os.path.basename(self.p3_file1_path)}")
            raw1 = pd.read_excel(self.p3_file1_path, header=None, dtype=str)

            header_row_idx = None
            prod_col_idx = 1  # 默认第2列
            for ri in range(min(5, len(raw1))):
                for ci, val in enumerate(raw1.iloc[ri].tolist()):
                    if '商品' in str(val) and '名' in str(val):
                        header_row_idx = ri
                        prod_col_idx = ci
                        break
                if header_row_idx is not None:
                    break

            if header_row_idx is None:
                header_row_idx = 0
                self._p3_log("  未找到'商品名称'表头，默认使用第2列")
            else:
                self._p3_log(
                    f"  表头行: 第{header_row_idx + 1}行，"
                    f"商品列: 第{prod_col_idx + 1}列")

            # 跳过表头行后的子标题行（NaN或非数据行）
            data_start = header_row_idx + 1
            while data_start < len(raw1):
                first = str(raw1.iloc[data_start, 0]).strip()
                prod  = str(raw1.iloc[data_start, prod_col_idx]).strip()
                if prod not in ('nan', 'NaN', '') and first not in ('nan', 'NaN', ''):
                    break
                data_start += 1

            product_rows = []
            for ri in range(data_start, len(raw1)):
                name = str(raw1.iloc[ri, prod_col_idx]).strip()
                if name not in ('nan', 'NaN', ''):
                    product_rows.append((name, ri + 1))  # 1-based for openpyxl

            self._p3_log(f"  文件1共 {len(product_rows)} 条商品数据")

            # ── 保存副本并用 openpyxl 上色 ───────────────────────────
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"比较订货结果_{timestamp}.xlsx"
            out_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            if not out_path:
                self._p3_log("操作已取消。")
                return

            shutil.copy2(self.p3_file1_path, out_path)

            wb = load_workbook(out_path)
            ws = wb.active
            order_col = ws.max_column + 1
            total_cols = order_col

            header_cell = ws.cell(row=header_row_idx + 1, column=order_col)
            header_cell.value = "订单量"
            src_header = ws.cell(row=header_row_idx + 1, column=order_col - 1)
            if src_header.has_style:
                header_cell._style = copy(src_header._style)

            green_fill = PatternFill(start_color="90EE90",
                                     end_color="90EE90", fill_type="solid")
            red_fill   = PatternFill(start_color="FF6B6B",
                                     end_color="FF6B6B", fill_type="solid")

            green_count = red_count = skip_count = 0
            for name, excel_row in product_rows:
                qty = order_map.get(name, 0.0)
                if qty == int(qty):
                    qty = int(qty)
                ws.cell(row=excel_row, column=order_col).value = qty
                if name in green_set:
                    fill = green_fill
                    green_count += 1
                elif name in zero_set:
                    skip_count += 1
                    continue  # 订单量=0，忽略不标色
                else:
                    fill = red_fill
                    red_count += 1
                for col in range(1, total_cols + 1):
                    ws.cell(row=excel_row, column=col).fill = fill

            wb.save(out_path)
            self._p3_log(f"✅ 已保存至: {out_path}")
            self._p3_log(
                f"   绿色（匹配）: {green_count} 条，"
                f"红色（未匹配）: {red_count} 条，"
                f"忽略（订单量=0）: {skip_count} 条")

            if messagebox.askyesno(
                "完成",
                f"比较完成！\n绿色（匹配）: {green_count} 条\n"
                f"红色（未匹配）: {red_count} 条\n"
                f"忽略（订单量=0）: {skip_count} 条\n\n是否打开结果文件？",
                parent=self.root
            ):
                try:
                    os.startfile(out_path)
                except Exception:
                    pass

        except Exception as e:
            self._p3_log(f"❌ 错误: {e}")
            messagebox.showerror("处理错误", str(e), parent=self.root)
        finally:
            self.p3_compare_btn.config(state=tk.NORMAL, text="开始比较")

    def _setup_page4(self):
        """创建其它功能页面的内容（预留扩展区）"""
        title_frame = tk.Frame(self.page4)
        title_frame.pack(pady=30, fill=tk.X)

        tk.Label(title_frame, text="🛠 其它功能",
                 font=("微软雅黑", 18, "bold")).pack()
        tk.Label(title_frame, text="此页面预留用于其它功能扩展",
                 font=("微软雅黑", 10), fg="gray").pack(pady=5)

        placeholder = tk.LabelFrame(self.page4, text="功能区",
                                     font=("微软雅黑", 10))
        placeholder.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        tk.Label(placeholder, text="功能开发中...",
                 font=("微软雅黑", 14), fg="#aaaaaa").pack(expand=True)

    def switch_page(self, page_idx):
        """切换显示的页面"""
        if page_idx == self.current_page:
            return
        pages = [self.page1, self.page2, self.page3, self.page4]
        pages[self.current_page].pack_forget()
        pages[page_idx].pack(fill=tk.BOTH, expand=True)
        self.current_page = page_idx
        self._update_tab_styles()

    def _update_tab_styles(self):
        """根据当前页更新标签按钮的高亮样式"""
        for i, btn in enumerate(self.tab_btns):
            if i == self.current_page:
                btn.config(bg="#ffffff", fg="#0066cc",
                           font=("微软雅黑", 10, "bold"))
            else:
                btn.config(bg="#f0f0f0", fg="#444444",
                           font=("微软雅黑", 10))

    def create_file_selection(self, label_text, file_index, parent=None):
        """创建文件选择组件"""
        if parent is None:
            parent = self.page1
        frame = tk.LabelFrame(parent, text=f"选择{label_text}", font=("微软雅黑", 10))
        frame.pack(pady=10, padx=20, fill=tk.X)
        
        # 按钮和标签框架
        inner_frame = tk.Frame(frame)
        inner_frame.pack(pady=10, padx=10, fill=tk.X)
        
        # 选择按钮
        if file_index == 0:
            command = self.select_file1
            var_attr = "btn_file1"
            label_attr = "label_file1"
        else:
            command = self.select_file2
            var_attr = "btn_file2"
            label_attr = "label_file2"
        
        button = ttk.Button(inner_frame, 
                           text=f"浏览...", 
                           command=command,
                           width=15)
        button.pack(side=tk.LEFT, padx=(0, 10))
        setattr(self, var_attr, button)
        
        # 文件路径标签
        label = tk.Label(inner_frame, 
                        text="未选择文件", 
                        font=("微软雅黑", 9),
                        fg="gray",
                        bg="#f0f0f0",
                        relief=tk.SUNKEN,
                        anchor=tk.W,
                        width=50)
        label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        setattr(self, label_attr, label)
    
    def select_file1(self):
        """选择第一个Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择第一个Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.file1_path = filename
            self.label_file1.config(text=os.path.basename(filename), fg="black")
            self.update_status(f"已选择文件1: {filename}")
            self.check_files_selected()
    
    def select_file2(self):
        """选择第二个Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择第二个Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.file2_path = filename
            self.label_file2.config(text=os.path.basename(filename), fg="black")
            self.update_status(f"已选择文件2: {filename}")
            self.check_files_selected()
    
    def check_files_selected(self):
        """检查是否已选择两个文件，如果是则启用合并按钮"""
        if getattr(self, 'p1_mode', None) and self.p1_mode.get() == "multi":
            return  # 多文件模式由 _p1_check_multi_ready 管理
        if self.file1_path and self.file2_path:
            self.merge_button.config(state=tk.NORMAL)
            self.status_bar.config(text="已选择两个文件，可以开始合并")
    
    def update_status(self, message):
        """在状态文本框添加新消息"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, f"> {message}\n")
        self.status_text.see(tk.END)  # 滚动到底部
        self.status_text.config(state=tk.DISABLED)
        self.root.update()  # 更新界面显示
    
    def apply_colors_to_excel(self, file_path, new_indices, merged_indices):
        """为Excel文件中的特定行应用颜色"""
        try:
            wb = load_workbook(file_path)
            ws = wb.active
            
            # 为新增项应用黄色
            for idx in new_indices:
                excel_row = idx + 2  # Excel行索引从1开始，表头占第1行
                for col in range(1, 4):  # 前三列
                    ws.cell(row=excel_row, column=col).fill = self.new_item_color
            
            # 为合并项应用绿色
            for idx in merged_indices:
                excel_row = idx + 2
                for col in range(1, 4):  # 前三列
                    ws.cell(row=excel_row, column=col).fill = self.merged_item_color
            
            wb.save(file_path)
            self.update_status("✅ 已为合并结果应用颜色标记")
            
        except Exception as e:
            self.update_status(f"⚠️ 应用颜色时出错: {str(e)}")
    
    def merge_files(self):
        """执行合并操作的核心函数"""
        if self.p1_mode.get() == "multi":
            self._p1_merge_multi()
            return

        if not (self.file1_path and self.file2_path):
            messagebox.showerror("错误", "请先选择两个Excel文件！")
            return
        
        # 禁用按钮，防止重复点击
        self.merge_button.config(state=tk.DISABLED, text="处理中...")
        self.status_bar.config(text="正在处理，请稍候...")
        self.root.update()
        
        try:
            # 读取文件
            self.update_status("正在读取文件...")
            df1 = pd.read_excel(self.file1_path)
            df2 = pd.read_excel(self.file2_path)
            
            # 检查必要列
            required_columns = ['商品', '盒码', '订单量']
            for df, name in [(df1, '文件1'), (df2, '文件2')]:
                missing = [col for col in required_columns if col not in df.columns]
                if missing:
                    raise ValueError(f"{name}缺少必要列: {missing}")
            
            self.update_status(f"文件1读取成功，共 {len(df1)} 行")
            self.update_status(f"文件2读取成功，共 {len(df2)} 行")
            
            # 处理逻辑
            result_df = df1.copy()
            df1_products = set(df1['商品'].astype(str).str.strip())
            df2_products = set(df2['商品'].astype(str).str.strip())
            
            # 找出新增商品
            new_products = df2_products - df1_products
            new_count = len(new_products)
            self.update_status(f"发现 {new_count} 个新增商品")
            
            # 记录新增项的索引
            new_indices = []
            if new_products:
                new_rows = df2[df2['商品'].astype(str).str.strip().isin(new_products)].copy()
                # 新增项的索引是原结果行数到原结果行数+新增行数-1
                start_idx = len(result_df)
                new_indices = list(range(start_idx, start_idx + len(new_rows)))
                result_df = pd.concat([result_df, new_rows], ignore_index=True, sort=False)
            
            # 处理重复商品，合并订单量
            common_products = df1_products.intersection(df2_products)
            common_count = len(common_products)
            self.update_status(f"发现 {common_count} 个共同商品，正在合并订单量...")
            
            # 创建商品到订单量的映射
            product_to_quantity = {}
            for _, row in df2.iterrows():
                product_name = str(row['商品']).strip()
                if product_name in common_products:
                    if product_name not in product_to_quantity:
                        product_to_quantity[product_name] = 0
                    product_to_quantity[product_name] += float(row['订单量'])
            
            # 更新结果中的订单量，并记录合并项的索引
            merged_indices = []
            merge_count = 0
            for idx, row in result_df.iterrows():
                product_name = str(row['商品']).strip()
                if product_name in product_to_quantity:
                    result_df.at[idx, '订单量'] = float(row['订单量']) + product_to_quantity[product_name]
                    merged_indices.append(idx)
                    merge_count += 1
            
            # 生成输出文件名（带时间戳）
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"合并结果_带颜色_{timestamp}.xlsx"
            
            # 让用户选择保存位置
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=output_filename,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not output_path:  # 用户取消了保存
                self.update_status("操作已取消")
                self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
                self.status_bar.config(text="操作取消")
                return
            
            # 保存结果
            result_df[['商品', '盒码', '订单量']].to_excel(output_path, index=False)
            
            # 为Excel文件添加颜色
            self.update_status("正在为合并结果添加颜色标记...")
            self.apply_colors_to_excel(output_path, new_indices, merged_indices)
            
            self.update_status("✅ 处理完成！")
            self.update_status(f"   总记录数: {len(result_df)} 行")
            self.update_status(f"   - 新增项（黄色）: {new_count} 行")
            self.update_status(f"   - 合并项（绿色）: {merge_count} 行")
            self.update_status(f"   文件已保存至: {output_path}")
            
            # 恢复按钮状态
            self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
            self.status_bar.config(text=f"处理完成！结果保存为: {os.path.basename(output_path)}")
            
            # 询问是否打开文件
            if messagebox.askyesno("完成", 
                                  f"合并完成！\n\n"
                                  f"总记录数: {len(result_df)} 行\n"
                                  f"新增项（黄色）: {new_count} 行\n"
                                  f"合并项（绿色）: {merge_count} 行\n\n"
                                  f"是否要打开结果文件？"):
                try:
                    os.startfile(output_path)
                except:
                    messagebox.showinfo("打开文件", f"文件已保存到:\n{output_path}")
            
        except FileNotFoundError as e:
            error_msg = f"文件未找到: {str(e)}\n请检查文件路径是否正确。"
            self.update_status(f"❌ 错误: {error_msg}")
            messagebox.showerror("文件错误", error_msg)
        except PermissionError as e:
            error_msg = f"文件被占用或无权限: {str(e)}\n请关闭已打开的Excel文件。"
            self.update_status(f"❌ 错误: {error_msg}")
            messagebox.showerror("权限错误", error_msg)
        except ValueError as e:
            error_msg = f"文件格式错误: {str(e)}\n请确保Excel文件包含'商品'、'盒码'、'订单量'三列。"
            self.update_status(f"❌ 错误: {error_msg}")
            messagebox.showerror("格式错误", error_msg)
        except Exception as e:
            error_msg = f"处理过程中发生未知错误:\n{str(e)}"
            self.update_status(f"❌ 错误: {error_msg}")
            messagebox.showerror("处理错误", error_msg)
            
            # 恢复按钮状态
            self.merge_button.config(state=tk.NORMAL, text="开始合并文件")
            self.status_bar.config(text="处理出错，请检查文件格式")
    
    def on_closing(self):
        """关闭窗口时的确认"""
        if messagebox.askokcancel("退出", "确定要退出程序吗？"):
            self.root.destroy()

    # ===== 自动更新相关方法 =====
    # 设计说明：
    #   运行中的 EXE 文件被 Windows 锁定，无法直接自我替换，也无需额外的 update.exe。
    #   方案：主程序从 GitHub 下载新 EXE 到系统临时目录，然后生成一个临时批处理脚本，
    #         该脚本等待主进程 PID 退出后，将新 EXE 覆盖旧 EXE，并重新启动。

    def check_for_updates(self):
        """点击检查更新按钮：在后台线程中获取并解析 version.json"""
        self.update_btn.config(state=tk.DISABLED, text="检查中...")
        self.update_version_label.config(
            text=f"当前版本: {CURRENT_VERSION}  正在检查更新..."
        )
        threading.Thread(target=self._check_update_thread, daemon=True).start()

    def _check_update_thread(self):
        """后台线程：根据 UPDATE_SOURCE 选择更新源，获取 version.json 并比对版本号"""
        try:
            headers = {"User-Agent": f"ExcelMergerApp/{CURRENT_VERSION}"}

            if UPDATE_SOURCE == "gitee":
                # ===== Gitee 更新源 =====
                version_url = f"{GITEE_RAW_BASE}/version.json"
                ver_req = urllib.request.Request(version_url, headers=headers)
                with urllib.request.urlopen(ver_req, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                # Gitee 版本使用 gitee_exe_url 字段，回退到 exe_url
                exe_url = str(data.get("gitee_exe_url") or data.get("exe_url", "")).strip()
                source_label = "Gitee"
            else:
                # ===== GitHub 更新源（默认） =====
                # 第一步：通过 GitHub API 获取仓库默认分支名（main / master / 其他）
                repo_req = urllib.request.Request(GITHUB_API_REPO_URL, headers=headers)
                with urllib.request.urlopen(repo_req, timeout=10) as resp:
                    repo_info = json.loads(resp.read().decode("utf-8"))
                default_branch = repo_info.get("default_branch", "main")

                # 第二步：用真实的分支名拼接 raw URL 读取 version.json
                version_url = (
                    f"https://raw.githubusercontent.com/{GITHUB_REPO}"
                    f"/{default_branch}/version.json"
                )
                ver_req = urllib.request.Request(version_url, headers=headers)
                with urllib.request.urlopen(ver_req, timeout=10) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                exe_url = str(data.get("exe_url", "")).strip()
                source_label = "GitHub"

            latest_version = str(data.get("version", "")).strip()
            release_notes = str(data.get("release_notes", "")).strip()

            if not latest_version:
                raise ValueError("version.json 中缺少 version 字段")

            def _ver(s):
                """将 'v2.1' 解析为可比较的元组 (2, 1)"""
                try:
                    return tuple(int(x) for x in s.lstrip("v").split("."))
                except Exception:
                    return (0,)

            if _ver(latest_version) > _ver(CURRENT_VERSION):
                self.root.after(0, lambda: self._on_update_available(
                    latest_version, exe_url, release_notes, source_label))
            else:
                self.root.after(0, lambda: self._on_no_update(source_label))

        except urllib.error.URLError as e:
            msg = str(e)
            self.root.after(0, lambda: self._on_update_error(f"网络连接失败: {msg}"))
        except Exception as e:
            msg = str(e)
            self.root.after(0, lambda: self._on_update_error(msg))

    def _on_update_available(self, new_version, exe_url, release_notes, source_label="GitHub"):
        """发现新版本：弹窗提示用户是否下载"""
        self.update_btn.config(state=tk.NORMAL, text="🔄 检查更新")
        self.update_version_label.config(
            text=f"当前版本: {CURRENT_VERSION}  ⬆ 发现新版本: {new_version} （{source_label}）"
        )
        notes_section = f"\n\n更新内容:\n{release_notes}" if release_notes else ""
        msg = (f"发现新版本 {new_version}！\n"
               f"当前版本: {CURRENT_VERSION}\n"
               f"下载源: {source_label}"
               f"{notes_section}\n\n"
               f"是否立即下载并安装更新？\n"
               f"（程序将自动关闭，安装完成后自动重启）")
        if messagebox.askyesno("发现新版本", msg):
            if exe_url:
                self._start_download(exe_url, new_version)
            else:
                messagebox.showinfo(
                    "更新提示",
                    f"version.json 中未提供 {'gitee_exe_url' if UPDATE_SOURCE == 'gitee' else 'exe_url'}，请前往 {source_label} 手动下载"
                )

    def _on_no_update(self, source_label="GitHub"):
        """已是最新版本"""
        self.update_btn.config(state=tk.NORMAL, text="🔄 检查更新")
        self.update_version_label.config(
            text=f"当前版本: {CURRENT_VERSION}  ✅ 已是最新版本"
        )
        messagebox.showinfo("检查更新", f"当前已是最新版本！（源: {source_label}）")

    def _on_update_error(self, error_msg):
        """检查或下载出错时恢复界面并提示"""
        self.update_btn.config(state=tk.NORMAL, text="🔄 检查更新")
        self.update_version_label.config(
            text=f"当前版本: {CURRENT_VERSION}  ⚠ 检查失败"
        )
        messagebox.showerror("检查更新失败", f"无法检查更新:\n{error_msg}")

    def _start_download(self, exe_url, new_version):
        """启动后台下载线程"""
        self.update_btn.config(state=tk.DISABLED, text="下载中...")
        self.update_version_label.config(
            text=f"当前版本: {CURRENT_VERSION}  正在下载 {new_version} 0%..."
        )
        threading.Thread(
            target=self._download_thread,
            args=(exe_url, new_version),
            daemon=True
        ).start()

    def _http_get_follow_redirects(self, url, headers, timeout=120):
        """
        使用自定义重定向处理器下载文件，确保每次重定向都保留完整请求头。
        urllib 默认在重定向时会丢弃自定义头（如 Accept），导致 GitHub CDN
        返回 HTML 而非 EXE，此方法通过覆盖 redirect_request 解决该问题。
        """
        original_headers = dict(headers)

        class _KeepHeadersRedirectHandler(urllib.request.HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, resp_headers, newurl):
                new_req = super().redirect_request(
                    req, fp, code, msg, resp_headers, newurl)
                if new_req is not None:
                    for key, val in original_headers.items():
                        new_req.add_unredirected_header(key, val)
                return new_req

        opener = urllib.request.build_opener(_KeepHeadersRedirectHandler())
        req = urllib.request.Request(url, headers=headers)
        resp = opener.open(req, timeout=timeout)
        return resp, resp.url

    def _download_thread(self, exe_url, new_version):
        """后台线程：分块下载新 EXE，然后生成替换脚本"""
        import tempfile

        # 兼容 PyInstaller 打包的 EXE 和直接运行的 .py
        if getattr(sys, "frozen", False):
            current_exe = sys.executable  # 打包后为当前 .exe 的完整路径
        else:
            current_exe = os.path.abspath(__file__)

        tmp_dir = tempfile.gettempdir()
        new_exe_path = os.path.join(tmp_dir, "excel_merger_new.exe")
        bat_path = os.path.join(tmp_dir, "excel_merger_update.bat")

        try:
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                "Accept": "application/octet-stream, */*",
            }
            # 手动跟踪重定向，保证每一跳都发送完整请求头
            resp, final_url = self._http_get_follow_redirects(exe_url, headers)

            content_type = resp.headers.get("Content-Type", "")
            if "text/html" in content_type:
                raise ValueError(
                    f"下载到的是 HTML 页面而非 EXE\n"
                    f"最终请求地址: {final_url}\n"
                    f"请确认 Release 为公开仓库且 exe_url 指向文件本身"
                )

            total = int(resp.headers.get("Content-Length", 0) or 0)
            downloaded = 0
            chunk_size = 65536  # 64 KB
            with open(new_exe_path, "wb") as f:
                while True:
                    chunk = resp.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = int(downloaded * 100 / total)
                        self.root.after(0, lambda p=pct: self.update_version_label.config(
                            text=f"当前版本: {CURRENT_VERSION}  下载中 {p}%..."
                        ))

            # 校验：有效的 Windows PE 可执行文件以 MZ 开头
            with open(new_exe_path, "rb") as f:
                magic = f.read(2)
            if magic != b"MZ":
                raise ValueError(
                    f"下载的文件不是有效的 EXE（文件头: {magic.hex()}，非 MZ）\n"
                    f"文件大小: {downloaded} 字节，最终地址: {final_url}"
                )

            # 生成等待主进程退出后执行替换并重启的批处理脚本
            # 注意：路径中若含空格，move 命令已用双引号包裹
            pid = os.getpid()
            bat_lines = [
                "@echo off",
                ":waitloop",
                f'tasklist /fi "PID eq {pid}" 2>nul | findstr /I "{pid}" >nul 2>&1',
                "if not errorlevel 1 (",
                "    timeout /t 1 /nobreak >nul",
                "    goto waitloop",
                ")",
                f'move /Y "{new_exe_path}" "{current_exe}"',
                f'start "" "{current_exe}"',
                "(goto) 2>nul & del \"%~f0\"",
            ]
            bat_content = "\r\n".join(bat_lines) + "\r\n"
            # 使用 GBK 编码，确保中文 Windows 的 cmd.exe 能正常执行
            with open(bat_path, "w", encoding="gbk", errors="replace") as f:
                f.write(bat_content)

            self.root.after(0, lambda: self._launch_updater_and_exit(bat_path, new_version))

        except Exception as e:
            # 清理临时文件，避免留下损坏的文件
            for p in (new_exe_path, bat_path):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
            msg = str(e)
            # 404 时给出明确指引
            if "404" in msg or "Not Found" in msg.lower():
                msg = (f"下载地址不存在 (404)，请检查：\n"
                       f"1. Release 标签名是否正确（如 v2.1）\n"
                       f"2. 上传的 EXE 文件名是否与 exe_url 一致\n"
                       f"3. 仓库是否为公开仓库\n\n"
                       f"当前 exe_url:\n{exe_url}")
            self.root.after(0, lambda: self._on_update_error(f"下载失败: {msg}"))

    def _launch_updater_and_exit(self, bat_path, new_version):
        """下载完成：确认后以隐藏窗口启动批处理脚本，然后退出主程序"""
        self.update_btn.config(state=tk.NORMAL, text="🔄 检查更新")
        if not messagebox.askyesno(
            "下载完成",
            f"新版本 {new_version} 下载完毕！\n\n"
            f"点击「是」程序将立即关闭并自动完成安装，安装后自动重新启动。\n"
            f"点击「否」取消本次安装（下载文件已保存，下次可重新触发）。"
        ):
            self.update_version_label.config(
                text=f"当前版本: {CURRENT_VERSION}  已取消安装"
            )
            return

        # 以隐藏控制台窗口的方式启动批处理，避免黑色命令行窗口弹出
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        subprocess.Popen(
            ["cmd.exe", "/C", bat_path],
            startupinfo=startupinfo,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
        )
        self.root.destroy()

def main():
    """主函数"""
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()

if __name__ == "__main__":
    # 测试用的模拟数据创建函数
    def create_test_files():
        """创建测试用的Excel文件（仅用于演示）"""
        import pandas as pd
        
        # 测试数据1
        data1 = {
            '商品': ['商品A', '商品B', '商品C', '商品D'],
            '盒码': ['001', '002', '003', '004'],
            '订单量': [10, 20, 15, 5]
        }
        
        # 测试数据2
        data2 = {
            '商品': ['商品C', '商品D', '商品E', '商品F'],
            '盒码': ['003', '004', '005', '006'],
            '订单量': [5, 10, 8, 12]
        }
        
        df1 = pd.DataFrame(data1)
        df2 = pd.DataFrame(data2)
        
        df1.to_excel("测试文件1.xlsx", index=False)
        df2.to_excel("测试文件2.xlsx", index=False)
        print("测试文件已创建: 测试文件1.xlsx, 测试文件2.xlsx")
    
    # 运行主程序
    main()

