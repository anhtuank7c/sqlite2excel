import tkinter as tk
from tkinter import ttk, filedialog
import threading
import os
from tkinter import font


class DBConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Xuất excel từ Database Sqlite")
        self.root.geometry("600x500")

        # Load Inter font
        _ = os.path.join(os.path.dirname(__file__), "fonts")
        try:
            # Load fonts for tkinter
            font.families()  # Initialize font system
            font.Font(root=root, font=("Inter", 10)).actual()  # Load default Inter font
            default_font = ("Inter", 10)
            title_font = ("Inter", 24, "bold")
            subtitle_font = ("Inter", 12)
        except:
            # Fallback to system fonts if Inter is not available
            default_font = ("Helvetica", 10)
            title_font = ("Helvetica", 24, "bold")
            subtitle_font = ("Helvetica", 12)

        # Set default font
        self.root.option_add("*Font", default_font)

        # Configure style for ttk widgets
        style = ttk.Style()
        style.configure(".", font=default_font)

        # Title and Subtitle
        title = ttk.Label(root, text="Xuất excel từ Database Sqlite", font=title_font)
        title.pack(pady=10)

        subtitle = ttk.Label(
            root,
            text="Ứng dụng giúp bạn xuất file excel từ database sqlite",
            font=subtitle_font,
        )
        subtitle.pack(pady=5)

        # Configure style for ttk widgets
        style = ttk.Style()
        style.configure(".", font=("Inter", 10))

        # Input file selection
        input_frame = ttk.Frame(root)
        input_frame.pack(fill="x", padx=20, pady=10)

        self.input_path = tk.StringVar()
        ttk.Label(input_frame, text="Database:").pack(side="left")
        ttk.Entry(input_frame, textvariable=self.input_path).pack(
            side="left", fill="x", expand=True, padx=5
        )
        ttk.Button(input_frame, text="Chọn", command=self.browse_input).pack(
            side="right"
        )

        # Output directory selection
        output_frame = ttk.Frame(root)
        output_frame.pack(fill="x", padx=20, pady=10)

        self.output_path = tk.StringVar()
        ttk.Label(output_frame, text="Thư mục xuất file:").pack(side="left")
        ttk.Entry(output_frame, textvariable=self.output_path).pack(
            side="left", fill="x", expand=True, padx=5
        )
        ttk.Button(output_frame, text="Chọn", command=self.browse_output).pack(
            side="right"
        )

        # List frame
        list_frame = ttk.LabelFrame(root, text="Danh sách các bảng")
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Listbox
        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.listbox.yview
        )
        scrollbar.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # Add/Remove buttons
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(side="right", fill="y", padx=5)

        ttk.Button(button_frame, text="Thêm", command=self.add_item).pack(pady=2)
        ttk.Button(button_frame, text="Xóa", command=self.remove_item).pack(pady=2)
        ttk.Button(
            button_frame,
            text="Đồng bộ",
            command=lambda: self.load_tables_from_db(self.input_path.get()),
        ).pack(pady=2)

        # Progress bar
        self.progress = ttk.Progressbar(root, length=400, mode="determinate")
        self.progress.pack(pady=10)

        # Execute and Cancel buttons
        button_frame = ttk.Frame(root)
        button_frame.pack(pady=10)

        self.execute_btn = ttk.Button(
            button_frame, text="Xuất file", command=self.execute, state="disabled"
        )
        self.execute_btn.pack(side="left", padx=5)

        self.cancel_btn = ttk.Button(
            button_frame, text="Hủy", command=self.cancel, state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=5)

        # Variable to track processing state
        self.is_processing = False
        self.should_cancel = False

        # Bind validation to path changes and list changes
        self.input_path.trace_add("write", self.validate_inputs)
        self.output_path.trace_add("write", self.validate_inputs)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self.validate_inputs())

    def browse_input(self):
        file_path = filedialog.askopenfilename(filetypes=[("Database file", "*.db")])
        if file_path:
            self.input_path.set(file_path)
            self.load_tables_from_db(file_path)

    def load_tables_from_db(self, db_path):
        if not db_path:
            return

        try:
            import sqlite3

            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            # Get list of tables
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()

            # Store current selection
            current_tables = list(self.listbox.get(0, tk.END))

            # Clear and repopulate list
            self.listbox.delete(0, tk.END)

            # Add tables to list
            seen = set()
            for table in tables:
                if table[0] not in seen:
                    self.listbox.insert(tk.END, table[0])
                    seen.add(table[0])

            # Add back any manually added tables that weren't in DB
            for table in current_tables:
                if table not in seen:
                    self.listbox.insert(tk.END, table)
                    seen.add(table)

            conn.close()
            self.validate_inputs()

        except Exception as e:
            tk.messagebox.showerror("Error", f"Lỗi truy vấn cơ sở dữ liệu: {str(e)}")

    def browse_output(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_path.set(dir_path)

    def add_item(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Thêm bảng")
        dialog.geometry("300x100")

        entry = ttk.Entry(dialog)
        entry.pack(padx=20, pady=10)

        def submit():
            text = entry.get().strip()
            if text:
                # Check if table already exists in list
                existing_tables = list(self.listbox.get(0, tk.END))
                if text in existing_tables:
                    tk.messagebox.showerror("Error", "Bảng này đã tồn tại")
                    return
                self.listbox.insert(tk.END, text)
                self.validate_inputs()
                dialog.destroy()

        ttk.Button(dialog, text="Thêm", command=submit).pack()

    def remove_item(self):
        selection = self.listbox.curselection()
        if selection:
            self.listbox.delete(selection)

    def validate_inputs(self, *args):
        if (
            self.input_path.get()
            and self.output_path.get()
            and self.listbox.size() > 0
            and not self.is_processing
        ):
            self.execute_btn["state"] = "normal"
        else:
            self.execute_btn["state"] = "disabled"

    def execute(self):
        if not self.input_path.get() or not self.output_path.get():
            tk.messagebox.showerror(
                "Error", "Bạn cần trỏ file database, thư mục xuất file"
            )
            return

        # Reset progress bar and flags
        self.progress["value"] = 0
        self.is_processing = True
        self.should_cancel = False

        # Update button states
        self.execute_btn["state"] = "disabled"
        self.cancel_btn["state"] = "normal"

        # Start processing in a separate thread
        thread = threading.Thread(target=self.process_data)
        thread.start()

    def cancel(self):
        self.should_cancel = True
        self.cancel_btn["state"] = "disabled"

    def process_data(self):
        try:
            import sqlite3
            import pandas as pd
            from datetime import datetime

            # Get paths and tables
            db_path = self.input_path.get()
            output_dir = self.output_path.get()
            tables = list(self.listbox.get(0, tk.END))

            # Reset processing state after completion
            def reset_state():
                self.is_processing = False
                self.cancel_btn["state"] = "disabled"
                self.validate_inputs()

            if not tables:
                tk.messagebox.showerror(
                    "Error", "Bạn cần thêm ít nhất một bảng vào danh sách"
                )
                return

            # Connect to database
            conn = sqlite3.connect(db_path)

            # Calculate progress step
            progress_step = 100 / len(tables)
            current_progress = 0

            # Generate timestamp for filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            # Process each table
            for table in tables:
                if self.should_cancel:
                    break

                try:
                    # Read table from database
                    df = pd.read_sql_query(f"SELECT * FROM {table}", conn)

                    # Create Excel writer
                    excel_path = f"{output_dir}/{table}_{timestamp}.xlsx"
                    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                        df.to_excel(writer, sheet_name=table, index=False)

                    # Update progress
                    current_progress += progress_step
                    self.progress["value"] = min(current_progress, 100)
                    self.root.update_idletasks()

                except Exception as e:
                    tk.messagebox.showerror(
                        "Error", f"Lỗi truy vấn bảng {table}: {str(e)}"
                    )

            conn.close()
            if not self.should_cancel:
                self.progress["value"] = 100
                self.root.update_idletasks()
                tk.messagebox.showinfo("Success", "Xuất file thành công")
            else:
                tk.messagebox.showinfo("Cancelled", "Dừng xuất file")
            reset_state()

        except Exception as e:
            tk.messagebox.showerror("Error", f"Xử lý lỗi: {str(e)}")
            self.progress["value"] = 0
            self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app = DBConverterApp(root)
    root.mainloop()
