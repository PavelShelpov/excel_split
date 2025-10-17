import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
import os
from core.processing import process_file
from excel_utils.analysis import get_all_sheets_headers
from excel_utils.filtering import select_categories_sequentially

logger = logging.getLogger('excel_splitter')

class ExcelSplitterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Splitter")
        self.root.geometry("800x600")
        
        # Инициализируем переменные
        self.source_file = tk.StringVar()
        self.destination_folder = tk.StringVar()
        self.columns = []
        self.selected_columns = []
        self.filters = {}
        self.valid_sheets = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        """Создает элементы интерфейса"""
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Фрейм для выбора файлов
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        # Выбор исходного файла
        ttk.Label(file_frame, text="Source Excel File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.source_file, width=70).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="Browse", command=self.browse_source).grid(row=0, column=2, padx=5, pady=2)
        
        # Выбор целевой папки
        ttk.Label(file_frame, text="Destination Folder:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.destination_folder, width=70).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="Browse", command=self.browse_destination).grid(row=1, column=2, padx=5, pady=2)
        
        # Фрейм для колонок
        columns_frame = ttk.LabelFrame(main_frame, text="Filter Columns", padding="10")
        columns_frame.pack(fill=tk.X, pady=5)
        
        # Кнопка анализа файла
        ttk.Button(columns_frame, text="Analyze File", command=self.analyze_file).pack(side=tk.LEFT, padx=5)
        
        # Список колонок
        columns_label = ttk.Label(columns_frame, text="Available columns:")
        columns_label.pack(side=tk.LEFT, padx=5)
        
        # Фрейм для логов
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Текстовое поле для логов
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand= scrollbar.set)
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Run", command=self.run_processing).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Clear", command=self.clear_log).pack(side=tk.RIGHT, padx=5)
    
    def log(self, message):
        """Записывает сообщение в лог-окно"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
    
    def browse_source(self):
        """Открывает диалог выбора исходного файла"""
        file_path = filedialog.askopenfilename(
            title="Select Source Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xlsm")]
        )
        if file_path:
            self.source_file.set(file_path)
    
    def browse_destination(self):
        """Открывает диалог выбора целевой папки"""
        folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if folder_path:
            self.destination_folder.set(folder_path)
    
    def analyze_file(self):
        """Анализирует файл и показывает доступные колонки"""
        source = self.source_file.get()
        if not source or not os.path.exists(source):
            messagebox.showerror("Error", "Please select a valid source file")
            return
            
        try:
            # Анализируем файл
            self.log(f"Analyzing file: {source}")
            sheet_headers = get_all_sheets_headers(source)
            self.valid_sheets = {sheet: data for sheet, data in sheet_headers.items() if data[0] is not None}
            
            if not self.valid_sheets:
                self.log("Error: No headers found in any sheet")
                return
                
            # Поиск пересечения заголовков
            all_headers = [set(headers) for headers, _ in self.valid_sheets.values()]
            common_headers = set.intersection(*all_headers) if all_headers else set()
            
            if not common_headers:
                self.log("Warning: No common headers found between sheets")
                return
                
            # Отображаем колонки
            self.columns = list(common_headers)
            self.log(f"Found {len(self.columns)} common columns:")
            for i, col in enumerate(self.columns, 1):
                self.log(f"  {i}. {col}")
                
        except Exception as e:
            self.log(f"Error analyzing file: {str(e)}")
    
    def run_processing(self):
        """Запускает обработку файла"""
        source = self.source_file.get()
        destination = self.destination_folder.get()
        
        if not source or not os.path.exists(source):
            messagebox.showerror("Error", "Please select a valid source file")
            return
            
        if not destination or not os.path.isdir(destination):
            messagebox.showerror("Error", "Please select a valid destination folder")
            return
        
        try:
            # Имитируем процесс обработки
            self.log("Starting file processing...")
            
            # Временно сохраняем стандартный вывод
            import sys
            original_stdout = sys.stdout
            from io import StringIO
            sys.stdout = StringIO()
            
            # Запускаем основную функцию обработки
            success = process_file()
            
            # Получаем вывод
            output = sys.stdout.getvalue()
            sys.stdout = original_stdout
            
            # Отображаем результаты
            self.log("Processing completed!")
            if success:
                self.log("Operation was successful")
            else:
                self.log("Operation failed")
            
            # Показываем сообщение об успехе
            if success:
                messagebox.showinfo("Success", "File processing completed successfully!")
            else:
                messagebox.showerror("Error", "File processing failed")
                
        except Exception as e:
            self.log(f"Unexpected error: {str(e)}")
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
    
    def clear_log(self):
        """Очищает лог-окно"""
        self.log_text.delete(1.0, tk.END)

def launch_gui():
    """Запускает графический интерфейс"""
    root = tk.Tk()
    app = ExcelSplitterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    launch_gui()