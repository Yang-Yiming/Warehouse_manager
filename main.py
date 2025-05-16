import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import datetime
import openpyxl
import os
import json

class WarehouseManager:
    def __init__(self, root):
        self.root = root
        self.root.title('仓库物资管理系统')
        self.data = []  # 存储物资信息的列表
        
        # 初始化路径
        self.init_paths()
        
        # 加载配置
        self.load_config()
        
        # 加载数据
        self.load_data()
        
        # 创建界面
        self.create_widgets()
        
    def init_paths(self):
        """初始化路径设置"""
        self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
        self.data_file = os.path.join(self.data_dir, 'warehouse_data.json')
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        
        # 确保目录存在
        for directory in [self.data_dir, self.output_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                
    def load_config(self):
        """加载配置文件"""
        self.organizations = []
        self.labels = []
        self.operators = []
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.organizations = config.get('organization', {}).get('val', [])
                    self.labels = config.get('labels', {}).get('val', [])
                    self.operators = config.get('operators', {}).get('val', [])
            except Exception as e:
                messagebox.showerror('配置加载错误', f'无法加载配置: {str(e)}')
                
    def load_data(self):
        """从文件加载数据"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                # 确保旧数据兼容性：为没有操作记录的数据添加空操作记录
                for item in self.data:
                    if '操作记录' not in item:
                        item['操作记录'] = []
            except Exception as e:
                messagebox.showerror('数据加载错误', f'无法加载数据: {str(e)}')
                
    def save_data(self):
        """保存数据到文件"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror('数据保存错误', f'无法保存数据: {str(e)}')

    def create_widgets(self):
        """创建界面组件"""
        self.create_search_panel()
        self.create_button_panel()
        self.create_table()
        self.create_quantity_panel()
        self.update_table()
    
    def create_search_panel(self):
        """创建搜索面板"""
        search_frame = tk.Frame(self.root)
        search_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(search_frame, text='搜索:').pack(side=tk.LEFT)
        
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<KeyRelease>', lambda e: self.update_table())

    def create_button_panel(self):
        """创建按钮面板"""
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(btn_frame, text='添加物资', command=self.add_item).pack(side=tk.LEFT)
        tk.Button(btn_frame, text='出库', command=self.remove_item).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text='导出为Excel', command=self.export_excel).pack(side=tk.LEFT, padx=5)

    def create_table(self):
        """创建数据表格"""
        columns = ('编号', '名称', '所属组织', '数量', '入库日期', '标签', '最近操作者')
        self.tree = ttk.Treeview(self.root, columns=columns, show='headings')
        
        # 配置列
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            if col in ['标签', '最近操作者']:
                self.tree.column(col, width=150)  # 标签列宽设置大一些
            else:
                self.tree.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        # 绑定双击事件以查看详细操作记录
        self.tree.bind('<Double-1>', self.show_operation_history)

    def show_operation_history(self, event):
        """显示物品的操作历史记录"""
        selected = self.tree.selection()
        if not selected:
            return
            
        idx = self.tree.index(selected[0])
        item = self.data[idx]
        
        # 创建操作历史窗口
        history_win = tk.Toplevel(self.root)
        history_win.title(f"{item['名称']} - 操作历史")
        history_win.geometry('600x400')
        
        # 创建表格显示操作历史
        columns = ('操作时间', '操作者', '操作类型', '操作数量', '操作后数量')
        history_tree = ttk.Treeview(history_win, columns=columns, show='headings')
        
        for col in columns:
            history_tree.heading(col, text=col)
            history_tree.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(history_win, orient="vertical", command=history_tree.yview)
        history_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        history_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 填充操作历史数据
        if '操作记录' in item and item['操作记录']:
            for record in reversed(item['操作记录']):  # 从新到旧显示
                history_tree.insert('', tk.END, values=(
                    record['时间'], record['操作者'], record['操作类型'], 
                    record['操作数量'], record['操作后数量']
                ))
        else:
            messagebox.showinfo('提示', '该物品暂无操作记录')
            history_win.destroy()

    def create_quantity_panel(self):
        """创建数量操作面板"""
        self.qty_frame = tk.Frame(self.root)
        self.qty_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(self.qty_frame, text="数量操作:").pack(side=tk.LEFT)
        
        # 减号按钮
        self.minus_btn = tk.Button(self.qty_frame, text="-", width=3, 
                            command=lambda: self.change_quantity(-1))
        self.minus_btn.pack(side=tk.LEFT, padx=5)
        
        # 数量编辑框
        self.qty_var = tk.StringVar()
        self.qty_entry = tk.Entry(self.qty_frame, textvariable=self.qty_var, width=8)
        self.qty_entry.pack(side=tk.LEFT, padx=5)
        self.qty_entry.bind('<Return>', lambda e: self.update_quantity_from_entry())
        self.qty_entry.bind('<FocusOut>', lambda e: self.update_quantity_from_entry())
        
        # 加号按钮
        self.plus_btn = tk.Button(self.qty_frame, text="+", width=3,
                           command=lambda: self.change_quantity(1))
        self.plus_btn.pack(side=tk.LEFT, padx=5)
        
        # 操作者输入框
        tk.Label(self.qty_frame, text="操作者:").pack(side=tk.LEFT, padx=(20, 5))
        self.operator_var = tk.StringVar()
        
        # 使用下拉列表选择操作者，但也允许手动输入
        self.operator_entry = ttk.Combobox(self.qty_frame, textvariable=self.operator_var, 
                                       values=self.operators, width=10)
        self.operator_entry.pack(side=tk.LEFT, padx=5)
        
        # 初始禁用控件
        self.disable_quantity_controls()

    def on_tree_select(self, event):
        """处理表格选择事件"""
        selected = self.tree.selection()
        if selected:
            idx = self.tree.index(selected[0])
            if idx < len(self.data):
                item = self.data[idx]
                self.enable_quantity_controls(item['数量'])
            else:
                self.disable_quantity_controls()
        else:
            self.disable_quantity_controls()
    
    def disable_quantity_controls(self):
        """禁用数量操作控件"""
        for widget in [self.minus_btn, self.qty_entry, self.plus_btn, self.operator_entry]:
            widget.config(state=tk.DISABLED)
        self.qty_var.set("")
        self.operator_var.set("")
    
    def enable_quantity_controls(self, qty):
        """启用数量操作控件并设置当前值"""
        for widget in [self.minus_btn, self.qty_entry, self.plus_btn, self.operator_entry]:
            widget.config(state=tk.NORMAL)
        self.qty_var.set(str(qty))
    
    def validate_operator(self):
        """验证操作者是否填写"""
        operator = self.operator_var.get().strip()
        if not operator:
            messagebox.showwarning('提示', '请输入操作者姓名')
            return False
        return True
    
    def change_quantity(self, change_amount):
        """增加或减少物品数量"""
        selected = self.tree.selection()
        if not selected:
            return
            
        # 验证操作者是否填写
        if not self.validate_operator():
            return
            
        idx = self.tree.index(selected[0])
        current_qty = self.data[idx]['数量']
        new_qty = current_qty + change_amount
        
        # 确保数量不小于0
        if new_qty < 0:
            messagebox.showwarning('提示', '库存数量不能小于0')
            return
        
        # 记录操作信息
        operation_type = "增加" if change_amount > 0 else "减少"
        self.add_operation_record(idx, self.operator_var.get(), operation_type, abs(change_amount), new_qty)
            
        self.data[idx]['数量'] = new_qty
        self.save_data()
        self.update_table()
        
        # 重新选择当前行
        self.reselect_row(idx, new_qty)
    
    def update_quantity_from_entry(self):
        """从输入框更新数量"""
        selected = self.tree.selection()
        if not selected:
            return
            
        # 验证操作者是否填写
        if not self.validate_operator():
            return
            
        try:
            new_qty = int(self.qty_var.get())
            if new_qty < 0:
                messagebox.showwarning('提示', '库存数量不能小于0')
                return
                
            idx = self.tree.index(selected[0])
            current_qty = self.data[idx]['数量']
            change_amount = new_qty - current_qty
            
            if change_amount == 0:  # 数量没有变化
                return
                
            # 记录操作信息
            operation_type = "设置为" 
            self.add_operation_record(idx, self.operator_var.get(), operation_type, new_qty, new_qty)
                
            self.data[idx]['数量'] = new_qty
            self.save_data()
            self.update_table()
            
            # 重新选择当前行
            self.reselect_row(idx, new_qty)
        except ValueError:
            messagebox.showwarning('提示', '请输入有效的数字')
    
    def add_operation_record(self, idx, operator, operation_type, operation_amount, new_qty):
        """添加操作记录"""
        if '操作记录' not in self.data[idx]:
            self.data[idx]['操作记录'] = []
            
        record = {
            '时间': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            '操作者': operator,
            '操作类型': operation_type,
            '操作数量': operation_amount,
            '操作后数量': new_qty
        }
        
        self.data[idx]['操作记录'].append(record)
    
    def reselect_row(self, idx, qty):
        """重新选择表格中的行"""
        if idx < len(self.tree.get_children()):
            item_id = self.tree.get_children()[idx]
            self.tree.selection_set(item_id)
            self.enable_quantity_controls(qty)

    def update_table(self):
        """更新表格数据显示"""
        search = self.search_var.get().lower()
        
        # 清空表格
        for row in self.tree.get_children():
            self.tree.delete(row)
            
        # 重新填充表格
        for item in self.data:
            # 兼容没有标签字段的旧数据
            tags = item.get('标签', [])
            tags_str = ", ".join(tags)
            
            # 获取最近操作者列表
            recent_operators = []
            if '操作记录' in item and item['操作记录']:
                # 从最新的记录开始，获取不重复的操作者
                seen_operators = set()
                for record in reversed(item['操作记录']):
                    operator = record['操作者']
                    if operator not in seen_operators:
                        recent_operators.append(operator)
                        seen_operators.add(operator)
                    if len(recent_operators) >= 3:  # 只显示最近3个不同的操作者
                        break
            
            recent_operators_str = ", ".join(recent_operators)
            
            # 检查是否符合搜索条件
            searchable_fields = [str(item['编号']), item['名称'], item['所属组织'], 
                               str(item['数量']), item['入库日期'], tags_str, recent_operators_str]
            if search and not any(search in field.lower() for field in searchable_fields):
                continue
                
            self.tree.insert('', tk.END, values=(
                item['编号'], item['名称'], item['所属组织'], 
                item['数量'], item['入库日期'], tags_str, recent_operators_str))
    
    def generate_new_id(self):
        """生成新的两位数字编号（01-99）"""
        used_ids = {item['编号'] for item in self.data}
        for i in range(1, 100):
            new_id = f"{i:02d}"  # 格式化为两位数字
            if new_id not in used_ids:
                return new_id
        return None  # 如果所有编号都被使用了

    def add_item(self):
        """添加新物资"""
        self.open_add_item_dialog()
    
    def open_add_item_dialog(self):
        """打开添加物资对话框"""
        win = tk.Toplevel(self.root)
        win.title('添加物资')
        win.geometry('400x550')  # 增加窗口大小以容纳操作者输入
        
        # 创建基本字段
        tk.Label(win, text='编号').grid(row=0, column=0, padx=5, pady=5, sticky='w')
        entry_id = tk.Entry(win)
        entry_id.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        
        tk.Label(win, text='名称').grid(row=1, column=0, padx=5, pady=5, sticky='w')
        entry_name = tk.Entry(win)
        entry_name.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        
        tk.Label(win, text='所属组织').grid(row=2, column=0, padx=5, pady=5, sticky='w')
        # 使用下拉列表选择组织
        org_var = tk.StringVar()
        if self.organizations:
            org_var.set(self.organizations[0])
        org_dropdown = ttk.Combobox(win, textvariable=org_var, values=self.organizations, state="readonly")
        org_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        
        tk.Label(win, text='数量').grid(row=3, column=0, padx=5, pady=5, sticky='w')
        entry_count = tk.Entry(win)
        entry_count.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        
        tk.Label(win, text='入库日期(YYYY-MM-DD)').grid(row=4, column=0, padx=5, pady=5, sticky='w')
        entry_date = tk.Entry(win)
        entry_date.grid(row=4, column=1, padx=5, pady=5, sticky='ew')
        
        # 添加操作者输入
        tk.Label(win, text='操作者').grid(row=5, column=0, padx=5, pady=5, sticky='w')
        operator_var = tk.StringVar()
        operator_entry = ttk.Combobox(win, textvariable=operator_var, values=self.operators)
        operator_entry.grid(row=5, column=1, padx=5, pady=5, sticky='ew')
        
        # 标签选择
        tk.Label(win, text='物品标签 (可多选)').grid(row=6, column=0, columnspan=2, padx=5, pady=5, sticky='w')
        
        # 创建标签复选框
        label_frame = tk.Frame(win)
        label_frame.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky='ew')
        
        label_vars = []
        for i, label in enumerate(self.labels):
            var = tk.BooleanVar()
            cb = tk.Checkbutton(label_frame, text=label, variable=var)
            cb.grid(row=i//2, column=i%2, sticky='w')
            label_vars.append((label, var))
        
        # 自动生成编号并填入
        new_id = self.generate_new_id()
        if new_id:
            entry_id.insert(0, new_id)
            
        # 默认填入当前日期
        entry_date.insert(0, datetime.date.today().strftime('%Y-%m-%d'))
        
        # 保存按钮
        tk.Button(win, text='保存', 
                 command=lambda: self.save_new_item(win, entry_id, entry_name, 
                                                   org_var, entry_count, entry_date, 
                                                   operator_var, label_vars)
                ).grid(row=8, column=0, columnspan=2, pady=10)
    
    def save_new_item(self, win, entry_id, entry_name, org_var, entry_count, entry_date, operator_var, label_vars):
        """保存新添加的物资"""
        try:
            item_id = entry_id.get()
            operator = operator_var.get().strip()
            
            # 验证编号是否为两位数字
            if not (item_id.isdigit() and len(item_id) == 2):
                raise ValueError('编号必须是两位数字(01-99)')
            
            # 检查编号是否重复
            if any(item['编号'] == item_id for item in self.data):
                raise ValueError('编号已存在，请使用其他编号')
            
            # 验证操作者是否填写
            if not operator:
                raise ValueError('请输入操作者姓名')
                
            # 获取选中的标签
            selected_labels = [label for label, var in label_vars if var.get()]
            
            try:
                count = int(entry_count.get())
                if count < 0:
                    raise ValueError('数量不能为负数')
            except:
                raise ValueError('请输入有效的数量')
            
            item = {
                '编号': item_id,
                '名称': entry_name.get(),
                '所属组织': org_var.get(),
                '数量': count,
                '入库日期': entry_date.get(),
                '标签': selected_labels,
                '操作记录': []
            }
            
            # 验证所有字段
            if not item['编号'] or not item['名称'] or not item['所属组织'] or not item['入库日期']:
                raise ValueError('编号、名称、所属组织和入库日期为必填项')
                
            # 验证日期格式
            datetime.datetime.strptime(item['入库日期'], '%Y-%m-%d')
            
            # 添加操作记录
            record = {
                '时间': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                '操作者': operator,
                '操作类型': '新增物品',
                '操作数量': count,
                '操作后数量': count
            }
            item['操作记录'].append(record)
            
            # 添加并保存
            self.data.append(item)
            self.save_data()
            self.update_table()
            win.destroy()
            
            # 如果操作者不在列表中，添加到配置
            if operator and operator not in self.operators:
                self.operators.append(operator)
                self.save_operators_to_config()
            
        except Exception as e:
            messagebox.showerror('错误', str(e))
    
    def save_operators_to_config(self):
        """保存操作者列表到配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 更新或添加操作者配置
                if 'operators' not in config:
                    config['operators'] = {
                        'description': '操作者列表',
                        'val': self.operators
                    }
                else:
                    config['operators']['val'] = self.operators
                
                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=4)
            except Exception as e:
                messagebox.showerror('配置保存错误', f'无法保存配置: {str(e)}')

    def remove_item(self):
        """出库物资"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning('提示', '请先选择要出库的物资')
            return
        
        # 验证操作者是否填写
        if not self.validate_operator():
            return
        
        idx = self.tree.index(selected[0])
        item = self.data[idx]
        
        # 创建出库方式选择对话框
        choice = messagebox.askyesnocancel('出库方式', '选择出库方式:\n是 - 部分出库\n否 - 完全出库\n取消 - 取消操作')
        
        if choice is None:  # 用户点击了取消
            return
        elif choice:  # 用户选择部分出库
            self.remove_item_partially(idx, item)
        else:  # 用户选择完全出库
            self.remove_item_completely(idx)
    
    def remove_item_partially(self, idx, item):
        """部分出库处理"""
        current_qty = item['数量']
        out_qty = simpledialog.askinteger('部分出库', 
                                         f'当前库存: {current_qty}\n请输入要出库的数量:', 
                                         minvalue=1, maxvalue=current_qty)
        
        if out_qty is None:  # 用户取消输入
            return
            
        if out_qty == current_qty:  # 如果出库数量等于当前库存，完全出库
            self.remove_item_completely(idx)
        else:  # 部分出库
            operator = self.operator_var.get()
            
            # 记录操作信息
            self.add_operation_record(idx, operator, "出库", out_qty, current_qty - out_qty)
            
            self.data[idx]['数量'] = current_qty - out_qty
            self.save_data()
            messagebox.showinfo('出库成功', f'已出库 {out_qty} 个 {item["名称"]}，剩余 {current_qty - out_qty} 个。')
            self.update_table()
            
    def remove_item_completely(self, idx):
        """完全出库处理"""
        if messagebox.askyesno('确认', '确定要将该物资完全出库吗？'):
            operator = self.operator_var.get()
            item = self.data[idx]
            
            # 记录到操作历史文件
            self.log_operation_to_file(item, operator, "完全出库", item['数量'], 0)
            
            del self.data[idx]
            self.save_data()
            self.update_table()
            messagebox.showinfo('出库成功', '物资已完全出库！')
            
    def log_operation_to_file(self, item, operator, operation, amount, new_qty):
        """记录操作到历史文件"""
        now = datetime.datetime.now()
        log_dir = os.path.join(self.data_dir, 'logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
            
        log_file = os.path.join(log_dir, f'operation_log_{now.strftime("%Y%m")}.txt')
        
        log_entry = (f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] "
                    f"操作者: {operator}, 物品: {item['名称']}(编号:{item['编号']}), "
                    f"操作: {operation}, 数量: {amount}, 剩余: {new_qty}\n")
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(log_entry)

    def sort_by(self, col, reverse):
        """按列排序表格数据"""
        col_map = {
            '编号': '编号', '名称': '名称', '所属组织': '所属组织', 
            '数量': '数量', '入库日期': '入库日期'
        }
        
        if col in col_map:
            self.data.sort(key=lambda x: x[col_map[col]], reverse=reverse)
            self.update_table()
            # 下次点击反向排序
            self.tree.heading(col, command=lambda: self.sort_by(col, not reverse))

    def export_excel(self):
        """导出数据为Excel文件"""
        if not self.data:
            messagebox.showinfo('提示', '没有数据可导出')
            return
            
        # 设置默认保存位置和文件名
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"仓库物资_{timestamp}"
        
        file_path = filedialog.asksaveasfilename(
            initialdir=self.output_dir,
            initialfile=default_filename,
            defaultextension='.xlsx', 
            filetypes=[('Excel文件', '*.xlsx')]
        )
        
        if not file_path:
            return
            
        # 创建Excel文件
        self.create_excel_file(file_path)
        
        # 导出操作记录到文本文件
        txt_file_path = os.path.splitext(file_path)[0] + "_操作记录.txt"
        self.export_operation_history(txt_file_path)
    
    def create_excel_file(self, file_path):
        """创建Excel文件"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "仓库物资"
        
        # 添加表头
        ws.append(['编号', '名称', '所属组织', '数量', '入库日期', '标签', '最近操作者'])
        
        for item in self.data:
            # 兼容没有标签字段的旧数据
            tags = item.get('标签', [])
            tags_str = ", ".join(tags)
            
            # 获取最近操作者
            recent_operators = []
            if '操作记录' in item and item['操作记录']:
                seen_operators = set()
                for record in reversed(item['操作记录']):
                    operator = record['操作者']
                    if operator not in seen_operators:
                        recent_operators.append(operator)
                        seen_operators.add(operator)
                    if len(recent_operators) >= 3:
                        break
                        
            recent_operators_str = ", ".join(recent_operators)
            
            ws.append([
                item['编号'], item['名称'], item['所属组织'], 
                item['数量'], item['入库日期'], tags_str, recent_operators_str
            ])
        
        # 添加操作记录工作表
        ws_history = wb.create_sheet(title="操作记录")
        ws_history.append(['编号', '名称', '时间', '操作者', '操作类型', '操作数量', '操作后数量'])
        
        for item in self.data:
            if '操作记录' in item and item['操作记录']:
                for record in item['操作记录']:
                    ws_history.append([
                        item['编号'], item['名称'], 
                        record['时间'], record['操作者'], record['操作类型'],
                        record['操作数量'], record['操作后数量']
                    ])
        
        wb.save(file_path)
        messagebox.showinfo('导出成功', f'数据已导出到 {file_path}')
    
    def export_operation_history(self, file_path):
        """导出操作记录到文本文件"""
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write("仓库物资操作记录\n")
            f.write("=" * 50 + "\n\n")
            
            for item in self.data:
                f.write(f"物品: {item['名称']} (编号: {item['编号']})\n")
                f.write("-" * 30 + "\n")
                
                if '操作记录' in item and item['操作记录']:
                    for record in reversed(item['操作记录']):  # 从新到旧
                        f.write(f"时间: {record['时间']}\n")
                        f.write(f"操作者: {record['操作者']}\n")
                        f.write(f"操作: {record['操作类型']} {record['操作数量']}\n")
                        f.write(f"操作后数量: {record['操作后数量']}\n")
                        f.write("-" * 20 + "\n")
                else:
                    f.write("暂无操作记录\n")
                
                f.write("\n")
        
        messagebox.showinfo('导出成功', f'操作记录已导出到 {file_path}')

if __name__ == '__main__':
    root = tk.Tk()
    app = WarehouseManager(root)
    root.mainloop()
