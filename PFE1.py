import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class ExcelAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Data Analyzer")
        self.root.geometry("600x800")
        self.root.configure(bg="#f0f8ff")

        self.style = ttk.Style()
        self.style.configure("TFrame", background="#e0ffff")
        self.style.configure("TLabel", background="#e0ffff", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10), padding=5)
        self.style.configure("TCombobox", font=("Arial", 10))

        self.file_path = tk.StringVar()

        self.original_data = None
        self.cleaned_data = None

        self.create_widgets()

    def create_widgets(self):
        # Main Frame
        main_frame = ttk.Frame(self.root, padding="15 15 15 15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File Upload Frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=10)

        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        upload_btn = ttk.Button(file_frame, text="Select Excel File", command=self.upload_excel)
        upload_btn.pack(side=tk.RIGHT)

        # Column Selection
        column_frame = ttk.Frame(main_frame)
        column_frame.pack(fill=tk.X, pady=10)

        ttk.Label(column_frame, text="Category Column:").grid(row=0, column=0, sticky="w")
        self.category_dropdown = ttk.Combobox(column_frame, state="disabled")
        self.category_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(column_frame, text="Numeric Column:").grid(row=1, column=0, sticky="w")
        self.numeric_dropdown = ttk.Combobox(column_frame, state="disabled")
        self.numeric_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Aggregation Method for Category
        ttk.Label(column_frame, text="Category Aggregation:").grid(row=2, column=0, sticky="w")
        self.category_aggregation_dropdown = ttk.Combobox(column_frame, values=["Count", "Unique"], state="disabled")
        self.category_aggregation_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Aggregation Method for Numeric Column
        ttk.Label(column_frame, text="Numeric Aggregation:").grid(row=3, column=0, sticky="w")
        self.numeric_aggregation_dropdown = ttk.Combobox(column_frame, values=["Sum", "Mean", "Count"], state="disabled")
        self.numeric_aggregation_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        column_frame.columnconfigure(1, weight=1)

        # Graph Type Selection
        graph_frame = ttk.Frame(main_frame)
        graph_frame.pack(fill=tk.X, pady=10)

        ttk.Label(graph_frame, text="Select Graph Type:").grid(row=0, column=0, sticky="w")
        self.graph_type_dropdown = ttk.Combobox(graph_frame, values=["Bar", "Line", "Pie"], state="disabled")
        self.graph_type_dropdown.grid(row=0, column=1, padx=5, sticky="ew")

        # Clean Data & Analyze Button
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)

        self.clean_var = tk.BooleanVar(value=True)
        clean_check = ttk.Checkbutton(action_frame, text="Clean Data", variable=self.clean_var)
        clean_check.pack(side=tk.LEFT)

        self.analyze_btn = ttk.Button(action_frame, text="Analyze & Visualize", command=self.analyze_data, state=tk.DISABLED)
        self.analyze_btn.pack(side=tk.RIGHT)

        # Matplotlib Figure and Canvas for Displaying Plot
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=main_frame)  # A tk.DrawingArea.
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, pady=15)

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
            self.category_dropdown['values'] = columns
            self.numeric_dropdown['values'] = columns
            self.category_aggregation_dropdown.set("Count")  # Default aggregation method
            self.numeric_aggregation_dropdown.set("Sum")  # Default aggregation method
            self.category_dropdown.config(state="readonly")
            self.numeric_dropdown.config(state="readonly")
            self.category_aggregation_dropdown.config(state="readonly")
            self.numeric_aggregation_dropdown.config(state="readonly")
            self.graph_type_dropdown.config(state="readonly")
            self.analyze_btn.config(state=tk.NORMAL)

            messagebox.showinfo("Success", "File uploaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {str(e)}")

    def analyze_data(self):
        try:
            category_col = self.category_dropdown.get()
            numeric_col = self.numeric_dropdown.get()
            graph_type = self.graph_type_dropdown.get()
            category_agg_method = self.category_aggregation_dropdown.get()
            numeric_agg_method = self.numeric_aggregation_dropdown.get()

            if not category_col or not numeric_col or not graph_type or not category_agg_method or not numeric_agg_method:
                messagebox.showwarning("Missing Input", "Please select all required options.")
                return

            self.cleaned_data = self.original_data.copy()
            if self.clean_var.get():
                self.cleaned_data = self.cleaned_data[[category_col, numeric_col]].dropna()

            # Category Aggregation
            if category_agg_method == "Count":
                category_aggregation = self.cleaned_data[category_col].value_counts()
            elif category_agg_method == "Unique":
                category_aggregation = self.cleaned_data[category_col].nunique()

            # Numeric Aggregation
            if numeric_agg_method == "Sum":
                numeric_aggregation = self.cleaned_data.groupby(category_col)[numeric_col].sum()
            elif numeric_agg_method == "Mean":
                numeric_aggregation = self.cleaned_data.groupby(category_col)[numeric_col].mean()
            elif numeric_agg_method == "Count":
                numeric_aggregation = self.cleaned_data.groupby(category_col)[numeric_col].count()

            category_aggregation = category_aggregation.sort_values(ascending=False)
            numeric_aggregation = numeric_aggregation.sort_values(ascending=False)

            # Clear previous plot
            self.ax.clear()

            # Plotting Graph
            if graph_type == "Bar":
                sns.barplot(x=numeric_aggregation.index, y=numeric_aggregation.values, palette='viridis', ax=self.ax)
                self.ax.set_title(f'{numeric_agg_method} of {numeric_col} by {category_col} - Bar Chart', fontsize=14)
            elif graph_type == "Line":
                self.ax.plot(numeric_aggregation.index, numeric_aggregation.values, marker='o')
                self.ax.set_title(f'{numeric_agg_method} of {numeric_col} by {category_col} - Line Chart', fontsize=14)
            elif graph_type == "Pie":
                self.ax.pie(numeric_aggregation.values, labels=numeric_aggregation.index, autopct='%1.1f%%', startangle=140)
                self.ax.set_title(f'{numeric_agg_method} of {numeric_col} by {category_col} - Pie Chart', fontsize=14)

            if graph_type != "Pie":
                self.ax.set_xlabel(category_col)
                self.ax.set_ylabel(f'{numeric_agg_method} of {numeric_col}')
                self.ax.tick_params(axis='x', rotation=45)

            # Draw the plot on canvas
            self.canvas.draw()

            # Save the summary CSV
            summary_df = pd.DataFrame({
                category_col: numeric_aggregation.index,
                f'{numeric_agg_method} of {numeric_col}': numeric_aggregation.values
            })
            summary_path = os.path.join(os.getcwd(), 'aggregated_summary.csv')
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
