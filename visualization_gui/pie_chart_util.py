import pandas as pd
import matplotlib.pyplot as plt
import os
from tkinter import messagebox

def generate_category_pie_chart(excel_file):
    df = pd.read_excel(excel_file, sheet_name=None, header=None)
    category_counts = {sheet_name: len(sheet_data) for sheet_name, sheet_data in df.items() if not sheet_data.empty}
    labels = list(category_counts.keys())[::-1]
    sizes = list(category_counts.values())[::-1]
    fig, ax = plt.subplots(figsize=(6, 6))
    patches, texts, pcts = ax.pie(
        sizes, labels=labels, autopct='%.1f%%',
        wedgeprops={'linewidth': 3.0, 'edgecolor': 'white'},
        textprops={'size': 'x-large'},
        startangle=90
    )
    for i, patch in enumerate(patches):
        texts[i].set_color(patch.get_facecolor())
    plt.setp(pcts, color='black')
    plt.setp(texts, fontweight=600)
    ax.set_title('Defect Categories', fontsize=18)
    plt.tight_layout()
    output_pie_chart = f"{os.path.splitext(excel_file)[0]}_pie_chart.png"
    plt.savefig(output_pie_chart)
    plt.close(fig)
    messagebox.showinfo("Pie Chart Saved", f"Pie chart saved as:\n{output_pie_chart}")
    os.startfile(output_pie_chart)