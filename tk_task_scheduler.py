#!/usr/bin/env python3
"""
Tk Task Scheduler GUI with multi-column job list (Name, Time, Command) and Test Run feature
Handles command paths with spaces when creating schtasks.
"""

import os, sys, json, subprocess, threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

APP_NAME = "TkTaskScheduler"
APPDATA = os.getenv('APPDATA') or os.path.join(Path.home(), 'AppData', 'Roaming')
DB_DIR = os.path.join(APPDATA, APP_NAME)
DB_PATH = os.path.join(DB_DIR, 'db.json')
WEEKDAY_MAP = ["MON","TUE","WED","THU","FRI","SAT","SUN"]
os.makedirs(DB_DIR, exist_ok=True)

def run_cmd(cmd):
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, shell=True, universal_newlines=True)
        return 0, out
    except subprocess.CalledProcessError as e:
        return e.returncode, e.output

def quote_command(cmd):
    # Wrap command properly to handle spaces
    if not cmd.strip():
        return cmd
    # If already quoted, leave it, else wrap with cmd.exe /c
    if ' ' in cmd and not (cmd.startswith('"') and cmd.endswith('"')):
        cmd = f'cmd.exe /c "{cmd}"'
    return cmd

def create_schtask(job):
    name, st, cmd = job['name'], job['time'], job['command']
    cmd = quote_command(cmd)
    if job['daily']:
        cmdline = f'schtasks /Create /TN "{name}" /TR "{cmd}" /SC DAILY /ST {st} /F'
    else:
        days = ','.join(job['days'])
        cmdline = f'schtasks /Create /TN "{name}" /TR "{cmd}" /SC WEEKLY /D {days} /ST {st} /F'
    return run_cmd(cmdline)

def delete_schtask(name):
    return run_cmd(f'schtasks /Delete /TN "{name}" /F')

def query_schtask(name):
    code, out = run_cmd(f'schtasks /Query /TN "{name}" /V /FO LIST')
    return (code == 0, out)

def load_db():
    try:
        with open(DB_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {'jobs': []}

def save_db(db):
    with open(DB_PATH, 'w', encoding='utf-8') as f:
        json.dump(db, f, indent=2)

class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_NAME)
        root.geometry('950x550')
        self.db = load_db()

        left = ttk.Frame(root)
        left.pack(side='left', fill='y', padx=6, pady=6)
        ttk.Label(left, text='Jobs').pack(anchor='w')

        columns = ('name', 'time', 'command')
        self.tree = ttk.Treeview(left, columns=columns, show='headings', height=20)
        self.tree.heading('name', text='Job Name')
        self.tree.heading('time', text='Time')
        self.tree.heading('command', text='Command')
        self.tree.column('name', width=140)
        self.tree.column('time', width=70)
        self.tree.column('command', width=250)
        self.tree.pack(fill='y', expand=True)
        self.tree.bind('<<TreeviewSelect>>', self.on_select_job)

        ttk.Button(left, text='Add New', command=self.clear_form).pack(fill='x', pady=3)
        ttk.Button(left, text='Delete', command=self.delete_selected).pack(fill='x', pady=3)
        ttk.Button(left, text='Test Run', command=self.test_run_selected).pack(fill='x', pady=3)

        right = ttk.Frame(root)
        right.pack(side='left', fill='both', expand=True, padx=6, pady=6)
        form = ttk.Frame(right)
        form.pack(fill='x')

        ttk.Label(form, text='Task Name').grid(row=0, column=0, sticky='w')
        self.name_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky='w')

        ttk.Label(form, text='Time (HH:MM)').grid(row=1, column=0, sticky='w')
        self.time_var = tk.StringVar(value='09:00')
        ttk.Entry(form, textvariable=self.time_var, width=10).grid(row=1, column=1, sticky='w')

        self.daily_var = tk.BooleanVar(value=True)
        self.daily_cb = ttk.Checkbutton(form, text='Daily', variable=self.daily_var, command=self.update_mode_state)
        self.daily_cb.grid(row=2, column=1, sticky='w')

        ttk.Label(form, text='Days of Week').grid(row=3, column=0, sticky='nw')
        days_frame = ttk.Frame(form)
        days_frame.grid(row=3, column=1, sticky='w')
        self.day_vars, self.day_cbs = [], []
        for i, d in enumerate(['Mon','Tue','Wed','Thu','Fri','Sat','Sun']):
            v = tk.BooleanVar()
            cb = ttk.Checkbutton(days_frame, text=d, variable=v, command=self.update_mode_state)
            cb.grid(row=i//4, column=i%4, sticky='w')
            self.day_vars.append(v)
            self.day_cbs.append(cb)

        ttk.Label(form, text='Command').grid(row=4, column=0, sticky='nw')
        self.cmd_text = ScrolledText(form, height=5, width=50)
        self.cmd_text.grid(row=4, column=1, sticky='w')

        ttk.Button(right, text='Save Task', command=self.save_task).pack(pady=6)
        self.log = ScrolledText(right, height=10)
        self.log.pack(fill='both', expand=True)

        self.refresh_jobs()
        self.update_mode_state()

    def update_mode_state(self):
        if self.daily_var.get():
            for cb in self.day_cbs:
                cb.state(['disabled'])
        else:
            any_day = any(v.get() for v in self.day_vars)
            if any_day:
                self.daily_cb.state(['disabled'])
            else:
                self.daily_cb.state(['!disabled'])
            for cb in self.day_cbs:
                cb.state(['!disabled'])

    def clear_form(self):
        self.name_var.set('')
        self.time_var.set('09:00')
        self.daily_var.set(True)
        for v in self.day_vars: v.set(False)
        self.cmd_text.delete('1.0', 'end')
        self.update_mode_state()

    def refresh_jobs(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for j in self.db['jobs']:
            self.tree.insert('', 'end', values=(j['name'], j['time'], j['command']))

    def on_select_job(self, e=None):
        sel = self.tree.selection()
        if not sel: return
        item = self.tree.item(sel[0])['values']
        name = item[0]
        job = next((j for j in self.db['jobs'] if j['name'] == name), None)
        if not job: return
        self.name_var.set(job['name'])
        self.time_var.set(job['time'])
        self.daily_var.set(job.get('daily', True))
        for i,v in enumerate(self.day_vars): v.set(WEEKDAY_MAP[i] in job.get('days', []))
        self.cmd_text.delete('1.0','end')
        self.cmd_text.insert('1.0', job['command'])
        self.update_mode_state()

    def save_task(self):
        name = self.name_var.get().strip()
        time = self.time_var.get().strip()
        cmd = self.cmd_text.get('1.0','end').strip()
        daily = self.daily_var.get()
        days = [WEEKDAY_MAP[i] for i,v in enumerate(self.day_vars) if v.get()]
        if not name or not cmd:
            messagebox.showerror(APP_NAME, 'Name and Command required')
            return
        if not daily and not days:
            messagebox.showerror(APP_NAME, 'Select at least one day or choose Daily')
            return
        job = {'name': name, 'time': time, 'command': cmd, 'daily': daily, 'days': days}
        self.db['jobs'] = [j for j in self.db['jobs'] if j['name'] != name] + [job]
        save_db(self.db)
        self.refresh_jobs()
        threading.Thread(target=lambda: self._create_task(job), daemon=True).start()

    def _create_task(self, job):
        exists, _ = query_schtask(job['name'])
        if exists: delete_schtask(job['name'])
        code, out = create_schtask(job)
        self.log.insert('end', out + '\n')
        self.log.see('end')

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel: return
        item = self.tree.item(sel[0])['values']
        name = item[0]
        job = next((j for j in self.db['jobs'] if j['name'] == name), None)
        if job and messagebox.askyesno('Confirm', f'Delete task {name}?'):
            delete_schtask(name)
            self.db['jobs'] = [j for j in self.db['jobs'] if j['name'] != name]
            save_db(self.db)
            self.refresh_jobs()

    def test_run_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, 'Select a job to test run')
            return
        item = self.tree.item(sel[0])['values']
        name = item[0]
        job = next((j for j in self.db['jobs'] if j['name'] == name), None)
        if not job:
            messagebox.showerror(APP_NAME, 'Job not found in database')
            return
        cmd = job['command']
        self.log.insert('end', f'Running test for {name}: {cmd}\n')
        threading.Thread(target=lambda: self._run_test(cmd), daemon=True).start()

    def _run_test(self, cmd):
        code, out = run_cmd(quote_command(cmd))
        self.log.insert('end', f'Exit {code}:\n{out}\n')
        self.log.see('end')

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == '__main__':
    main()
