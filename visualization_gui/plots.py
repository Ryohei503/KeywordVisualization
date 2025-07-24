import os
import sys
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.colors import Normalize
import seaborn as sns
from wordcloud import WordCloud
import pandas as pd
from tkinter import messagebox



def load_japanese_font():
    try:
        if getattr(sys, 'frozen', False):
            font_path = os.path.join(sys._MEIPASS, "src/ipaexg.ttf")
        else:
            font_path = os.path.join(os.path.dirname(__file__), "src/ipaexg.ttf")
        fm.fontManager.addfont(font_path)
        font_name = fm.FontProperties(fname=font_path).get_name()
        plt.rcParams['font.family'] = font_name
        return font_path
    except Exception as e:
        print(f"Font load failed: {e}")
        return None

def get_colordict(cmap_name, max_value, min_value=1):
    cmap = plt.get_cmap(cmap_name)
    norm = Normalize(vmin=min_value, vmax=max_value)
    return {i: cmap(norm(i)) for i in range(min_value, max_value + 1)}

def generate_graph(df_sorted, chunk_size, color_dict, top_n_label, sheet_name, output_path, sheet_names_global):
    df_sorted.columns = [col.lower() for col in df_sorted.columns]
    num_words = len(df_sorted)
    index_list = [
        [start, min(start + chunk_size, num_words)]
        for start in range(0, num_words, chunk_size)
    ]
    num_chunks = len(index_list)
    fig, axs = plt.subplots(1, num_chunks, figsize=(16, 8), facecolor='white', squeeze=False)
    max_count = df_sorted['count'].max()
    for col, idx in zip(range(num_chunks), index_list):
        df_chunk = df_sorted.iloc[idx[0]:idx[1]]
        labels = [f"{w}: {n}" for w, n in zip(df_chunk['word'], df_chunk['count'])]
        colors = [color_dict.get(i) for i in df_chunk['count']]
        x = list(df_chunk['count'])
        y = list(range(len(df_chunk)))
        # Assign y to hue and set legend=False to avoid deprecation warning
        sns.barplot(
            x=x, y=y, data=df_chunk, alpha=0.9, orient='h',
            ax=axs[0][col], hue=y, palette=colors, legend=False
        )
        axs[0][col].set_xlim(0, max_count + 1)
        axs[0][col].set_yticks(range(len(labels)))
        axs[0][col].set_yticklabels(labels, fontsize=12)
        for side in ['bottom', 'right', 'top', 'left']:
            axs[0][col].spines[side].set_color('white')
        ax = axs[0][col]
        ax.xaxis.set_ticks_position('top')
        ax.xaxis.set_label_position('top')
    # Title logic
    if len(sheet_names_global) == 1:
        fig.suptitle(f"Word Frequency Graph - {top_n_label}", fontsize=16)
    else:
        fig.suptitle(f"Word Frequency Graph - {sheet_name} - {top_n_label}", fontsize=16)
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()

def generate_wordcloud(df, n, font_path, output_path):
    df.columns = [col.lower() for col in df.columns]
    df = df.iloc[:n]
    freq_dict = pd.Series(df["count"].values, index=df["word"]).to_dict()
    wc = WordCloud(
        font_path=font_path,
        width=800,
        height=600,
        background_color="white",
        prefer_horizontal=1.0,
        colormap="tab20"
    )
    wc.generate_from_frequencies(freq_dict)
    plt.figure(figsize=(10, 8))
    plt.imshow(wc, interpolation="bilinear")
    plt.axis("off")
    plt.tight_layout(pad=0)
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()

def generate_bubble_chart(df_words, n, color_dict, output_path):
    import circlify
    df_words.columns = [col.lower() for col in df_words.columns]
    df_words = df_words.iloc[:n]
    labels = list(df_words['word'])
    counts = list(df_words['count'])
    labels.reverse()
    counts.reverse()
    circles = circlify.circlify(
        counts,
        show_enclosure=False,
        target_enclosure=circlify.Circle(x=0, y=0, r=1)
    )
    fig, ax = plt.subplots(figsize=(9, 9), facecolor='white')
    ax.axis('off')
    lim = max(
        max(abs(circle.x) + circle.r, abs(circle.y) + circle.r)
        for circle in circles
    )
    plt.xlim(-lim, lim)
    plt.ylim(-lim, lim)
    for circle, label, count in zip(circles, labels, counts):
        x, y, r = circle.x, circle.y, circle.r
        ax.add_patch(plt.Circle((x, y), r, alpha=0.9, color=color_dict.get(count)))
        plt.annotate(f"{label}\n{count}", (x, y), size=12, va='center', ha='center')
    plt.title("Keyword Frequency Bubble Chart")
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()

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