import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Data Analyzer")
        self.root.geometry("500x600")
        self.root.configure(bg="yellow")

        self.style = ttk.Style()
        self.style.configure("TFrame", background="orange")
        self.style.configure("TLabel", background="#add8e6")
        self.style.configure("TButton", background="#add8e6", font=("Arial", 10))

        self.main_frame = ttk.Frame(self.root, padding="20 20 20 20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.file_path = tk.StringVar()

        self.create_widgets()

        self.original_data = None
        self.cleaned_data = None

    def create_widgets(self):
        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        self.file_label = ttk.Label(file_frame, text="No file selected", font=("Arial", 10))
        self.file_label.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

        upload_btn = ttk.Button(file_frame, text="Select Excel File", command=self.upload_excel)
        upload_btn.pack(side=tk.RIGHT)

        column_frame = ttk.Frame(self.main_frame)
        column_frame.pack(fill=tk.X, pady=10)

        ttk.Label(column_frame, text="Select Category Column:").pack(side=tk.LEFT)
        self.column_dropdown = ttk.Combobox(column_frame, state="disabled")
        self.column_dropdown.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        graph_type_frame = ttk.Frame(self.main_frame)
        graph_type_frame.pack(fill=tk.X, pady=10)

        ttk.Label(graph_type_frame, text="Select Graph Type:").pack(side=tk.LEFT)
        self.graph_type_dropdown = ttk.Combobox(graph_type_frame, values=["Bar", "Line", "Pie"], state="disabled")
        self.graph_type_dropdown.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        analysis_frame = ttk.Frame(self.main_frame)
        analysis_frame.pack(fill=tk.X, pady=10)

        self.clean_var = tk.BooleanVar(value=True)
        clean_check = ttk.Checkbutton(analysis_frame, text="Clean Data", variable=self.clean_var)
        clean_check.pack(side=tk.LEFT)

        analyze_btn = ttk.Button(analysis_frame, text="Analyze & Visualize", command=self.analyze_data, state=tk.DISABLED)
        analyze_btn.pack(side=tk.RIGHT)
        self.analyze_btn = analyze_btn

    def upload_excel(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])

            if not file_path:
                return

            if file_path.endswith('.csv'):
                self.original_data = pd.read_csv(file_path)
            else:
                self.original_data = pd.read_excel(file_path)

            self.file_label.config(text=os.path.basename(file_path))

            columns = self.original_data.columns.tolist()
            self.column_dropdown['values'] = columns
            self.column_dropdown.config(state="readonly")
            self.graph_type_dropdown.config(state="readonly")

            self.analyze_btn.config(state=tk.NORMAL)

            messagebox.showinfo("Success", "File uploaded successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")

    def analyze_data(self):
        try:
            selected_column = self.column_dropdown.get()
            selected_graph_type = self.graph_type_dropdown.get()

            if not selected_column or not selected_graph_type:
                messagebox.showwarning("Warning", "Please select a category column and graph type")
                return

            self.cleaned_data = self.original_data.copy()

            if self.clean_var.get():
                self.cleaned_data.dropna(subset=[selected_column], inplace=True)
                self.cleaned_data = self.cleaned_data[self.cleaned_data[selected_column].notna()]

            category_counts = self.cleaned_data[selected_column].value_counts()

            plt.figure(figsize=(10, 6))
            if selected_graph_type == "Bar":
                sns.barplot(x=category_counts.index, y=category_counts.values, palette='coolwarm')
                plt.title(f'Distribution of {selected_column} - Bar Chart', fontsize=16)
            elif selected_graph_type == "Line":
                plt.plot(category_counts.index, category_counts.values, marker='o')
                plt.title(f'Distribution of {selected_column} - Line Chart', fontsize=16)
            elif selected_graph_type == "Pie":
                plt.pie(category_counts.values, labels=category_counts.index, autopct='%1.1f%%', startangle=140)
                plt.title(f'Distribution of {selected_column} - Pie Chart', fontsize=16)

            plt.xlabel(selected_column, fontsize=12)
            plt.ylabel('Count', fontsize=12)
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.show()

            summary_df = category_counts.reset_index()
            summary_df.columns = [selected_column, 'Count']
            summary_path = os.path.join(os.getcwd(), 'category_summary.csv')
            summary_df.to_csv(summary_path, index=False)
            messagebox.showinfo("Analysis Complete", f"Summary saved to {summary_path}")

        except Exception as e:
            messagebox.showerror("Analysis Error", str(e))


def main():
    root = tk.Tk()
    app = ExcelAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
