import os
import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, ttk
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import threading
import time

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class WebClickSchedulerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Web Click Scheduler")

        # 创建输入框和标签
        self.url_label = tk.Label(root, text="URL (打开网页):")
        self.url_label.grid(row=0, column=0, padx=10, pady=5)
        self.url_entry = tk.Entry(root, width=50)
        self.url_entry.grid(row=0, column=1, padx=10, pady=5)

        self.task_url_label = tk.Label(root, text="任务URL:")
        self.task_url_label.grid(row=1, column=0, padx=10, pady=5)
        self.task_url_entry = tk.Entry(root, width=50)
        self.task_url_entry.grid(row=1, column=1, padx=10, pady=5)

        self.xpath_label = tk.Label(root, text="XPath:")
        self.xpath_label.grid(row=2, column=0, padx=10, pady=5)
        self.xpath_entry = tk.Entry(root, width=50)
        self.xpath_entry.grid(row=2, column=1, padx=10, pady=5)

        self.hour_label = tk.Label(root, text="时:")
        self.hour_label.grid(row=3, column=0, padx=10, pady=5)
        self.hour_entry = tk.Entry(root, width=10)
        self.hour_entry.grid(row=3, column=1, padx=10, pady=5)

        self.minute_label = tk.Label(root, text="分:")
        self.minute_label.grid(row=4, column=0, padx=10, pady=5)
        self.minute_entry = tk.Entry(root, width=10)
        self.minute_entry.grid(row=4, column=1, padx=10, pady=5)

        self.second_label = tk.Label(root, text="秒:")
        self.second_label.grid(row=5, column=0, padx=10, pady=5)
        self.second_entry = tk.Entry(root, width=10)
        self.second_entry.grid(row=5, column=1, padx=10, pady=5)

        self.open_url_button = tk.Button(root, text="打开网页", command=self.open_url)
        self.open_url_button.grid(row=0, column=2, padx=10, pady=5)

        self.add_task_button = tk.Button(root, text="添加任务", command=self.add_task)
        self.add_task_button.grid(row=1, column=2, padx=10, pady=5)

        self.modify_task_button = tk.Button(root, text="修改任务", command=self.modify_selected_task)
        self.modify_task_button.grid(row=2, column=2, padx=10, pady=10)

        self.remove_task_button = tk.Button(root, text="删除任务", command=self.remove_selected_task)
        self.remove_task_button.grid(row=3, column=2, padx=10, pady=10)

        self.start_button = tk.Button(root, text="开始", command=self.start_clicks)
        self.start_button.grid(row=4, column=2, padx=10, pady=10)

        self.stop_button = tk.Button(root, text="停止", command=self.stop_clicks)
        self.stop_button.grid(row=5, column=2, padx=10, pady=10)

        # 创建任务显示框
        self.task_tree = ttk.Treeview(root, columns=("Sort", "Task URL", "XPath", "Delay"), show="headings")
        self.task_tree.heading("Sort", text="排序")
        self.task_tree.heading("Task URL", text="任务 URL")
        self.task_tree.heading("XPath", text="XPath")
        self.task_tree.heading("Delay", text="倒计时(HH:MM:SS)")
        self.task_tree.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        # 初始化任务列表和状态
        self.tasks = []
        self.scheduler_thread = None
        self.driver = None  # 初始化浏览器驱动为 None
        self.next_sort_value = 1  # 初始化下一个 sort 值

        # 配置网格布局
        root.grid_rowconfigure(6, weight=1)
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_columnconfigure(2, weight=1)

    def open_url(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("输入错误", "请输入有效的URL")
            return

        if self.driver is None:
            edge_options = Options()
            # edge_options.add_argument("--headless")  # 注释掉这一行以在前台查看浏览器行为
            edge_options.add_argument("--disable-gpu")
            edge_options.add_argument("--window-size=1920x1080")

            # 获取 msedgedriver.exe 的绝对路径
            driver_path = resource_path('msedgedriver.exe')  # 假设 msedgedriver.exe 在项目根目录
            service = Service(driver_path)
            self.driver = webdriver.Edge(service=service, options=edge_options)

        try:
            self.driver.get(url)
            print(f"访问页面: {url}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开URL: {e}")

    def add_task(self):
        task_url = self.task_url_entry.get().strip()
        xpath = self.xpath_entry.get().strip()
        hour = self.hour_entry.get().strip()
        minute = self.minute_entry.get().strip()
        second = self.second_entry.get().strip()

        if not task_url or not xpath or not hour or not minute or not second:
            messagebox.showwarning("输入错误", "请填写所有字段")
            return

        try:
            hour = int(hour)
            minute = int(minute)
            second = int(second)
            total_delay = hour * 3600 + minute * 60 + second
        except ValueError:
            messagebox.showwarning("输入错误", "小时、分钟、秒必须是整数")
            return

        sort = self.next_sort_value  # 自动赋值 sort 字段
        self.next_sort_value += 1  # 更新下一个 sort 值

        task = (task_url, xpath, total_delay, sort)
        self.tasks.append(task)
        self.update_task_tree()
        messagebox.showinfo("任务添加成功", f"已添加任务: Task URL={task_url}, XPath={xpath}, Delay={hour}:{minute}:{second}, Sort={sort}")

    def remove_selected_task(self):
        selected_item = self.task_tree.selection()
        if not selected_item:
            messagebox.showwarning("选择错误", "请选择一个任务")
            return

        item_values = self.task_tree.item(selected_item[0], "values")
        sort = int(item_values[0])
        task_url = item_values[1]
        xpath = item_values[2]
        delay = item_values[3]

        for task in self.tasks:
            if task[3] == sort and task[0] == task_url and task[1] == xpath and task[2] == self.parse_delay(delay):
                self.tasks.remove(task)
                self.update_task_tree()
                messagebox.showinfo("任务清除", f"已清除任务: Sort={sort}, Task URL={task_url}, XPath={xpath}, Delay={delay}")
                return

        messagebox.showwarning("选择错误", "未找到选定的任务")

    def modify_selected_task(self):
        selected_item = self.task_tree.selection()
        if not selected_item:
            messagebox.showwarning("选择错误", "请选择一个任务")
            return

        item_values = self.task_tree.item(selected_item[0], "values")
        sort = int(item_values[0])
        task_url = item_values[1]
        xpath = item_values[2]
        delay = item_values[3]

        # 解析现有的延迟时间
        hours, minutes, seconds = map(int, delay.split(':'))

        # 创建修改任务的弹窗
        modify_window = tk.Toplevel(self.root)
        modify_window.title("修改任务")

        # 创建输入框和标签
        tk.Label(modify_window, text="任务 URL:").grid(row=0, column=0, padx=10, pady=5)
        task_url_entry = tk.Entry(modify_window, width=50)
        task_url_entry.grid(row=0, column=1, padx=10, pady=5)
        task_url_entry.insert(0, task_url)

        tk.Label(modify_window, text="XPath:").grid(row=1, column=0, padx=10, pady=5)
        xpath_entry = tk.Entry(modify_window, width=50)
        xpath_entry.grid(row=1, column=1, padx=10, pady=5)
        xpath_entry.insert(0, xpath)

        tk.Label(modify_window, text="Hour:").grid(row=2, column=0, padx=10, pady=5)
        hour_entry = tk.Entry(modify_window, width=10)
        hour_entry.grid(row=2, column=1, padx=10, pady=5)
        hour_entry.insert(0, str(hours))

        tk.Label(modify_window, text="Minute:").grid(row=3, column=0, padx=10, pady=5)
        minute_entry = tk.Entry(modify_window, width=10)
        minute_entry.grid(row=3, column=1, padx=10, pady=5)
        minute_entry.insert(0, str(minutes))

        tk.Label(modify_window, text="Second:").grid(row=4, column=0, padx=10, pady=5)
        second_entry = tk.Entry(modify_window, width=10)
        second_entry.grid(row=4, column=1, padx=10, pady=5)
        second_entry.insert(0, str(seconds))

        tk.Label(modify_window, text="Sort (Priority):").grid(row=5, column=0, padx=10, pady=5)
        sort_entry = tk.Entry(modify_window, width=10)
        sort_entry.grid(row=5, column=1, padx=10, pady=5)
        sort_entry.insert(0, str(sort))

        def save_modifications():
            new_task_url = task_url_entry.get().strip()
            new_xpath = xpath_entry.get().strip()
            new_hour = hour_entry.get().strip()
            new_minute = minute_entry.get().strip()
            new_second = second_entry.get().strip()
            new_sort = sort_entry.get().strip()

            if not new_task_url or not new_xpath or not new_hour or not new_minute or not new_second or not new_sort:
                messagebox.showwarning("输入错误", "请填写所有字段")
                return

            try:
                new_hour = int(new_hour)
                new_minute = int(new_minute)
                new_second = int(new_second)
                new_sort = int(new_sort)
                new_total_delay = new_hour * 3600 + new_minute * 60 + new_second
            except ValueError:
                messagebox.showwarning("输入错误", "小时、分钟、秒和优先级必须是整数")
                return

            for task in self.tasks:
                if task[3] == sort and task[0] == task_url and task[1] == xpath and task[2] == self.parse_delay(delay):
                    self.tasks.remove(task)
                    new_task = (new_task_url, new_xpath, new_total_delay, new_sort)
                    self.tasks.append(new_task)
                    self.update_task_tree()
                    messagebox.showinfo("任务修改成功", f"已修改任务: Task URL={new_task_url}, XPath={new_xpath}, Delay={new_hour}:{new_minute}:{new_second}, Sort={new_sort}")
                    modify_window.destroy()
                    return

            messagebox.showwarning("选择错误", "未找到选定的任务")

        save_button = tk.Button(modify_window, text="保存修改", command=save_modifications)
        save_button.grid(row=6, column=0, columnspan=2, pady=10)

    def parse_delay(self, delay_str):
        hours, minutes, seconds = map(int, delay_str.split(':'))
        return hours * 3600 + minutes * 60 + seconds

    def start_clicks(self):
        if not self.tasks:
            messagebox.showwarning("无任务", "请先添加任务")
            return

        if self.driver is None:
            messagebox.showwarning("无浏览器实例", "请先打开一个网页")
            return

        if self.scheduler_thread and self.scheduler_thread.is_alive():
            messagebox.showwarning("任务正在运行", "已经有任务在运行中")
            return

        self.running = True  # 设置运行状态为 True
        self.scheduler_thread = threading.Thread(target=self.perform_clicks)
        self.scheduler_thread.start()

    def perform_clicks(self):
        while self.tasks and self.running:
            task_url, xpath, delay, sort = self.tasks.pop(0)
            try:
                self.driver.get(task_url)
                print(f"访问页面: {task_url}")
                # 显式等待元素出现并可点击
                element = WebDriverWait(self.driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                # 滚动到元素位置
                self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                # 检查元素是否可见且未被遮挡
                is_visible = self.driver.execute_script("""
                    var element = arguments[0];
                    var style = window.getComputedStyle(element);
                    return style.display !== 'none' && style.visibility !== 'hidden' && element.offsetWidth > 0 && element.offsetHeight > 0;
                """, element)
                if not is_visible:
                    print(f"元素 {xpath} 不可见或被遮挡")
                    continue
                # 尝试 Selenium 点击
                element.click()
                print(f"Selenium 点击元素 {xpath} 在页面 {task_url}.")
                time.sleep(delay)
            except Exception as e:
                print(f"Selenium 点击元素 {xpath} 失败: {e}")
                # 尝试使用 JavaScript 点击
                try:
                    self.driver.execute_script("arguments[0].click();", element)
                    print(f"JavaScript 点击元素 {xpath} 在页面 {task_url}.")
                    time.sleep(delay)
                except Exception as js_e:
                    print(f"JavaScript 点击元素 {xpath} 失败: {js_e}")
                    # 尝试使用 ActionChains 点击
                    try:
                        actions = ActionChains(self.driver)
                        actions.move_to_element(element).click().perform()
                        print(f"ActionChains 点击元素 {xpath} 在页面 {task_url}.")
                        time.sleep(delay)
                    except Exception as ac_e:
                        print(f"ActionChains 点击元素 {xpath} 失败: {ac_e}")
                        continue
            self.update_task_tree()

    def stop_clicks(self):
        self.running = False
        if self.scheduler_thread and self.scheduler_thread.is_alive():
            self.scheduler_thread.join()
        if self.driver:
            self.driver.quit()
            self.driver = None

    def update_task_tree(self):
        self.task_tree.delete(*self.task_tree.get_children())
        for task in sorted(self.tasks, key=lambda x: x[3]):
            task_url, xpath, delay, sort = task
            hours = delay // 3600
            minutes = (delay % 3600) // 60
            seconds = delay % 60
            self.task_tree.insert("", tk.END, values=(sort, task_url, xpath, f"{hours:02}:{minutes:02}:{seconds:02}"))

if __name__ == "__main__":
    root = tk.Tk()
    app = WebClickSchedulerGUI(root)
    app.running = True  # 初始化运行状态
    root.mainloop()