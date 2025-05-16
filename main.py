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
        
        # 加载数据
        self.load_data()
        
        # 创建界面
        self.create_widgets()
        
    def init_paths(self):
        """初始化路径设置"""
        self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
        self.data_file = os.path.join(self.data_dir, 'warehouse_data.json')
        
        # 确保目录存在
        for directory in [self.data_dir, self.output_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                
    def load_data(self):
        """从文件加载数据"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
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
        columns = ('编号', '名称', '所属组织', '数量', '入库日期')
        self.tree = ttk.Treeview(self.root, columns=columns, show='headings')
        
        # 配置列
        for col in columns:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            self.tree.column(col, width=100)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

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
        for widget in [self.minus_btn, self.qty_entry, self.plus_btn]:
            widget.config(state=tk.DISABLED)
        self.qty_var.set("")
    
    def enable_quantity_controls(self, qty):
        """启用数量操作控件并设置当前值"""
        for widget in [self.minus_btn, self.qty_entry, self.plus_btn]:
            widget.config(state=tk.NORMAL)
        self.qty_var.set(str(qty))
    
    def change_quantity(self, change_amount):
        """增加或减少物品数量"""
        selected = self.tree.selection()
        if not selected:
            return
            
        idx = self.tree.index(selected[0])
        current_qty = self.data[idx]['数量']
        new_qty = current_qty + change_amount
        
        # 确保数量不小于0
        if new_qty < 0:
            return
            
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
            
        try:
            new_qty = int(self.qty_var.get())
            if new_qty < 0:
                return
                
            idx = self.tree.index(selected[0])
            self.data[idx]['数量'] = new_qty
            self.save_data()
            self.update_table()
            
            # 重新选择当前行
            self.reselect_row(idx, new_qty)
        except ValueError:
            pass
    
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
            if any(search in str(v).lower() for v in item.values()):
                self.tree.insert('', tk.END, values=(
                    item['编号'], item['名称'], item['所属组织'], 
                    item['数量'], item['入库日期']))
    
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
        
        # 创建输入字段
        labels = ['编号', '名称', '所属组织', '数量', '入库日期(YYYY-MM-DD)']
        entries = []
        
        for i, label in enumerate(labels):
            tk.Label(win, text=label).grid(row=i, column=0, padx=5, pady=5)
            entry = tk.Entry(win)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries.append(entry)
            
        entry_id, entry_name, entry_org, entry_count, entry_date = entries
        
        # 自动生成编号并填入
        new_id = self.generate_new_id()
        if new_id:
            entry_id.insert(0, new_id)
            
        # 默认填入当前日期
        entry_date.insert(0, datetime.date.today().strftime('%Y-%m-%d'))
        
        # 保存按钮
        tk.Button(win, text='保存', 
                 command=lambda: self.save_new_item(win, entry_id, entry_name, 
                                                   entry_org, entry_count, entry_date)
                ).grid(row=len(labels), column=0, columnspan=2, pady=10)
    
    def save_new_item(self, win, entry_id, entry_name, entry_org, entry_count, entry_date):
        """保存新添加的物资"""
        try:
            item_id = entry_id.get()
            # 验证编号是否为两位数字
            if not (item_id.isdigit() and len(item_id) == 2):
                raise ValueError('编号必须是两位数字(01-99)')
            
            # 检查编号是否重复
            if any(item['编号'] == item_id for item in self.data):
                raise ValueError('编号已存在，请使用其他编号')
                
            item = {
                '编号': item_id,
                '名称': entry_name.get(),
                '所属组织': entry_org.get(),
                '数量': int(entry_count.get()),
                '入库日期': entry_date.get()
            }
            
            # 验证所有字段
            if not item['编号'] or not item['名称'] or not item['所属组织'] or not item['入库日期']:
                raise ValueError('所有字段均为必填')
                
            # 验证日期格式
            datetime.datetime.strptime(item['入库日期'], '%Y-%m-%d')
            
            # 添加并保存
            self.data.append(item)
            self.save_data()
            self.update_table()
            win.destroy()
            
        except Exception as e:
            messagebox.showerror('错误', str(e))

    def remove_item(self):
        """出库物资"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning('提示', '请先选择要出库的物资')
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
            self.data[idx]['数量'] = current_qty - out_qty
            self.save_data()
            messagebox.showinfo('出库成功', f'已出库 {out_qty} 个 {item["名称"]}，剩余 {current_qty - out_qty} 个。')
            self.update_table()
            
    def remove_item_completely(self, idx):
        """完全出库处理"""
        if messagebox.askyesno('确认', '确定要将该物资完全出库吗？'):
            del self.data[idx]
            self.save_data()
            self.update_table()
            messagebox.showinfo('出库成功', '物资已完全出库！')

    def sort_by(self, col, reverse):
        """按列排序表格数据"""
        col_map = {'编号': '编号', '名称': '名称', '所属组织': '所属组织', '数量': '数量', '入库日期': '入库日期'}
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
        default_filename = f"仓库物资_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
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
    
    def create_excel_file(self, file_path):
        """创建Excel文件"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['编号', '名称', '所属组织', '数量', '入库日期'])
        for item in self.data:
            ws.append([item['编号'], item['名称'], item['所属组织'], item['数量'], item['入库日期']])
        wb.save(file_path)
        messagebox.showinfo('导出成功', f'数据已导出到 {file_path}')

if __name__ == '__main__':
    root = tk.Tk()
    app = WarehouseManager(root)
    root.mainloop()
