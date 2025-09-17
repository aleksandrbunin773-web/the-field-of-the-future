import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Анализ данных мероприятий")
        self.root.geometry("1200x800")
        
        self.data = None
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        file_frame = ttk.LabelFrame(main_frame, text="Выбор файла данных", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(file_frame, text="Загрузить Excel файл", 
                  command=self.load_file).pack(side=tk.LEFT, padx=5)
        
        params_frame = ttk.LabelFrame(main_frame, text="Параметры анализа", padding="10")
        params_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        ttk.Label(params_frame, text="Ось X:").grid(row=0, column=0, sticky=tk.W)
        self.x_var = tk.StringVar(value="Момент подачи заявления")
        x_combo = ttk.Combobox(params_frame, textvariable=self.x_var, 
                              values=["Момент подачи заявления", "Наименование мероприятия", 
                                     "Наименование смены", "Грантовый конкурс"])
        x_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(params_frame, text="Ось Y:").grid(row=1, column=0, sticky=tk.W)
        self.y_var = tk.StringVar(value="Вовлеченность")
        y_combo = ttk.Combobox(params_frame, textvariable=self.y_var,
                              values=["Вовлеченность", "Количество участников", 
                                     "Активность", "Рейтинг"])
        y_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(params_frame, text="Наименование смены:").grid(row=2, column=0, sticky=tk.W)
        self.theme_var = tk.StringVar(value="все")
        theme_combo = ttk.Combobox(params_frame, textvariable=self.theme_var,
                                  values=["все"])
        theme_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Button(params_frame, text="Анализировать данные", 
                  command=self.analyze_data).grid(row=3, column=0, columnspan=2, pady=10)
        
        results_frame = ttk.LabelFrame(main_frame, text="Результаты анализа", padding="10")
        results_frame.grid(row=1, column=1, rowspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        ttk.Label(results_frame, text="Прогноз:").grid(row=0, column=0, sticky=tk.W)
        self.prediction_text = tk.Text(results_frame, height=5, width=40)
        self.prediction_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        ttk.Label(results_frame, text="Статистика:").grid(row=2, column=0, sticky=tk.W)
        self.stats_text = tk.Text(results_frame, height=8, width=40)
        self.stats_text.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
       
        graph_frame = ttk.LabelFrame(main_frame, text="Визуализация данных", padding="10")
        graph_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        help_frame = ttk.LabelFrame(main_frame, text="Помощь", padding="10")
        help_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        help_text = """
        X - Момент подачи заявления: Анализ по дате подачи заявлений
        X - Наименование мероприятия: Сравнение различных мероприятий
        X - Наименование смены: Анализ по сменам форума
        X - Грантовый конкурс: Сравнение участников грантового конкурса
        """
        ttk.Label(help_frame, text=help_text, justify=tk.LEFT).pack()
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.data = pd.read_excel(file_path)
                messagebox.showinfo("Успех", f"Файл загружен успешно!\nЗаписей: {len(self.data)}")
                self.preprocess_data()
                self.update_theme_combobox()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка загрузки файла: {str(e)}")
    
    def preprocess_data(self):
        """Предварительная обработка данных"""
        if self.data is None:
            return
            
        # Автоматическое определение названий колонок
        self.detect_column_names()
        
        # Преобразование дат
        if hasattr(self, 'timestamp_col') and self.timestamp_col:
            try:
                self.data[self.timestamp_col] = pd.to_datetime(self.data[self.timestamp_col], 
                                                             errors='coerce')
                # Извлекаем дату и время суток
                self.data['дата_подачи'] = self.data[self.timestamp_col].dt.date
                self.data['время_суток'] = self.data[self.timestamp_col].dt.hour
            except Exception as e:
                print(f"Ошибка преобразования даты: {e}")
        
        # Обработка грантового конкурса
        if hasattr(self, 'grant_col') and self.grant_col:
            # Создаем числовые значения для анализа
            self.data['грант_число'] = self.data[self.grant_col].apply(
                lambda x: 1 if isinstance(x, str) and 'да' in x.lower() else (
                    -1 if isinstance(x, str) and 'нет' in x.lower() else 0
                )
            )
    
    def detect_column_names(self):
        """Автоматическое определение названий колонок"""
        column_mapping = {
            'timestamp_col': ['момент подачи заявления', 'timestamp', 'дата', 'время', 'time', 'date'],
            'event_col': ['наименование мероприятия', 'мероприятие', 'event', 'title'],
            'shift_col': ['наименование смены', 'смена', 'shift', 'theme'],
            'grant_col': ['анкета_планируете ли вы участвовать', 'грант', 'grant', 'конкурс'],
            'engagement_col': ['вовлеченность', 'engagement', 'активность', 'activity', 'рейтинг', 'rating']
        }
        
        for attr_name, possible_names in column_mapping.items():
            found_col = None
            for col in self.data.columns:
                if any(name in str(col).lower() for name in possible_names):
                    found_col = col
                    break
            setattr(self, attr_name, found_col)
    
    def update_theme_combobox(self):
        """Обновление выпадающего списка с темами"""
        if hasattr(self, 'shift_col') and self.shift_col and self.shift_col in self.data.columns:
            themes = ["все"] + sorted(self.data[self.shift_col].dropna().unique().tolist())
            theme_combo = ttk.Combobox(self.root.nametowidget('.!frame.!labelframe.!frame.!combobox'),
                                     values=themes)
            theme_combo.set("все")
            self.theme_var.set("все")
    
    def analyze_data(self):
        if self.data is None:
            messagebox.showwarning("Предупреждение", "Сначала загрузите файл данных!")
            return
        
        try:
            # Фильтрация по тематике
            filtered_data = self.filter_by_theme(self.data)
            
            # Анализ данных
            self.create_plot(filtered_data)
            self.calculate_statistics(filtered_data)
            self.generate_prediction(filtered_data)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка анализа: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def filter_by_theme(self, data):
        theme = self.theme_var.get()
        if theme == "все" or not hasattr(self, 'shift_col') or not self.shift_col:
            return data
        
        return data[data[self.shift_col] == theme]
    
    def create_plot(self, data):
        self.ax.clear()
        
        x_selection = self.x_var.get()
        y_selection = self.y_var.get()
        
        # Определяем колонки для X и Y
        if x_selection == "Момент подачи заявления":
            x_col = 'дата_подачи' if 'дата_подачи' in data.columns else self.timestamp_col
        elif x_selection == "Наименование мероприятия":
            x_col = self.event_col
        elif x_selection == "Наименование смены":
            x_col = self.shift_col
        elif x_selection == "Грантовый конкурс":
            x_col = self.grant_col if hasattr(self, 'grant_col') else None
        
        if y_selection == "Вовлеченность":
            y_col = self.engagement_col if hasattr(self, 'engagement_col') else None
        else:
            y_col = None
        
        if x_col is None or y_col is None or x_col not in data.columns:
            self.ax.text(0.5, 0.5, "Недостаточно данных для построения графика", 
                        ha='center', va='center', transform=self.ax.transAxes)
            self.canvas.draw()
            return
        
        # Группировка данных
        if x_selection == "Момент подачи заявления" and pd.api.types.is_datetime64_any_dtype(data[x_col]):
            grouped = data.groupby(data[x_col].dt.date)[y_col].mean()
            self.ax.plot(grouped.index, grouped.values, 'o-', linewidth=2, markersize=6)
            self.ax.set_xlabel('Дата подачи заявления')
            plt.xticks(rotation=45)
        elif x_selection == "Грантовый конкурс":
            # Анализ грантового конкурса
            grant_data = data.groupby(self.grant_col)['грант_число'].count()
            self.ax.bar(grant_data.index, grant_data.values)
            self.ax.set_xlabel('Участие в грантовом конкурсе')
            plt.xticks(rotation=45)
        else:
            # Для категориальных данных
            grouped = data.groupby(x_col)[y_col].mean()
            x_pos = range(len(grouped))
            self.ax.bar(x_pos, grouped.values)
            self.ax.set_xticks(x_pos)
            self.ax.set_xticklabels(grouped.index, rotation=45, ha='right')
            self.ax.set_xlabel(x_selection)
        
        self.ax.set_ylabel(y_selection)
        self.ax.set_title(f'{y_selection} по {x_selection}')
        self.ax.grid(True, alpha=0.3)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def calculate_statistics(self, data):
        stats_text = ""
        
        # Статистика по грантовому конкурсу
        if hasattr(self, 'grant_col') and self.grant_col in data.columns:
            grant_yes = data[data['грант_число'] == 1].shape[0]
            grant_no = data[data['грант_число'] == -1].shape[0]
            total = grant_yes + grant_no
            
            stats_text += "=== СТАТИСТИКА ГРАНТОВОГО КОНКУРСА ===\n"
            stats_text += f"Планируют участвовать: {grant_yes} ({grant_yes/total*100:.1f}%)\n"
            stats_text += f"Не планируют участвовать: {grant_no} ({grant_no/total*100:.1f}%)\n\n"
            
            # Сравнение вовлеченности
            if hasattr(self, 'engagement_col') and self.engagement_col in data.columns:
                engagement_yes = data[data['грант_число'] == 1][self.engagement_col].mean()
                engagement_no = data[data['грант_число'] == -1][self.engagement_col].mean()
                
                stats_text += "=== ВОВЛЕЧЕННОСТЬ ===\n"
                stats_text += f"Участники гранта: {engagement_yes:.2f}\n"
                stats_text += f"Остальные: {engagement_no:.2f}\n"
                stats_text += f"Разница: {engagement_yes - engagement_no:.2f}\n\n"
        
        # Общая статистика
        if hasattr(self, 'engagement_col') and self.engagement_col in data.columns:
            stats_text += "=== ОБЩАЯ СТАТИСТИКА ===\n"
            stats_text += f"Средняя вовлеченность: {data[self.engagement_col].mean():.2f}\n"
            stats_text += f"Максимальная: {data[self.engagement_col].max():.2f}\n"
            stats_text += f"Минимальная: {data[self.engagement_col].min():.2f}\n"
            stats_text += f"Всего записей: {len(data)}\n"
        
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, stats_text)
    
    def generate_prediction(self, data):
        prediction = "АНАЛИЗ И ПРОГНОЗ\n\n"
        
        # Анализ грантового конкурса
        if hasattr(self, 'grant_col') and self.grant_col in data.columns:
            grant_ratio = data[data['грант_число'] == 1].shape[0] / len(data)
            
            prediction += "ГРАНТОВЫЙ КОНКУРС:\n"
            prediction += f"- Доля участников: {grant_ratio*100:.1f}%\n"
            
            if grant_ratio > 0.3:
                prediction += "- Высокий интерес к грантовой поддержке ✓\n"
            else:
                prediction += "- Низкий интерес к грантовой поддержке ⚠\n"
        
        # Анализ вовлеченности
        if hasattr(self, 'engagement_col') and self.engagement_col in data.columns:
            avg_engagement = data[self.engagement_col].mean()
            
            prediction += "\nВОВЛЕЧЕННОСТЬ:\n"
            prediction += f"- Средний показатель: {avg_engagement:.2f}\n"
            
            if avg_engagement > 0.7:
                prediction += "- Высокая вовлеченность участников ✓\n"
            elif avg_engagement > 0.4:
                prediction += "- Средняя вовлеченность участников ∼\n"
            else:
                prediction += "- Низкая вовлеченность участников ⚠\n"
        
        # Рекомендации
        prediction += "\nРЕКОМЕНДАЦИИ:\n"
        prediction += "- Провести дополнительные мотивационные мероприятия\n"
        prediction += "- Улучшить информирование о грантовых возможностях\n"
        prediction += "- Создать систему поддержки для участников грантов\n"
        prediction += "- Регулярно отслеживать вовлеченность участников"
        
        self.prediction_text.delete(1.0, tk.END)
        self.prediction_text.insert(1.0, prediction)

def main():
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()