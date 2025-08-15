import os
import sys
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.colors import Normalize
import seaborn as sns
from wordcloud import WordCloud
import pandas as pd


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

def generate_graph(df, chunk_size, color_dict, top_n_label, sheet_name, output_path, sheet_names_global):
    sns.set_style("darkgrid") # Apply Seaborn darkgrid style
    df.columns = [col.lower() for col in df.columns]
    num_words = len(df)
    index_list = [
        [start, min(start + chunk_size, num_words)]
        for start in range(0, num_words, chunk_size)
    ]
    num_chunks = len(index_list)
    fig, axs = plt.subplots(1, num_chunks, figsize=(16, 8), facecolor='white', squeeze=False)
    max_count = df['count'].max()
    # Load Japanese font for yticklabels
    try:
        font_path = load_japanese_font()
        font_prop = fm.FontProperties(fname=font_path) if font_path else None
    except Exception:
        font_prop = None
    for col, idx in zip(range(num_chunks), index_list):
        df_chunk = df.iloc[idx[0]:idx[1]]
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
        if font_prop:
            axs[0][col].set_yticklabels(labels, fontsize=12, fontproperties=font_prop)
        else:
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

def generate_wordcloud(df, n, font_path, output_path, sheet_name, top_n_label, sheet_names_global):
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
    # Title logic
    if len(sheet_names_global) == 1:
        plt.title(f"Word Frequency Graph - {top_n_label}", fontsize=16)
    else:
        plt.title(f"Word Frequency Graph - {sheet_name} - {top_n_label}", fontsize=16)
    plt.tight_layout(pad=0)
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()

def generate_bubble_chart(df_words, n, color_dict, output_path, sheet_name, top_n_label, sheet_names_global):
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
    if len(sheet_names_global) == 1:
        plt.title(f"Bubble Chart - {top_n_label}", fontsize=16)
    else:
        plt.title(f"Bubble Chart - {sheet_name} - {top_n_label}", fontsize=16)
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
    output_path = f"{os.path.splitext(excel_file)[0]}_pie_chart.png"
    plt.savefig(output_path)
    plt.close(fig)
    return output_path

def generate_category_box_plot(excel_file, selected_priority="All"):
    # Read data from Excel file
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    data = []
    sheet_order = list(excel_data.keys())

    # Accept list of priorities or 'All' or None
    priorities_filter = None
    show_priorities_in_title = False
    show_priorities_in_filename = False
    if isinstance(selected_priority, list):
        priorities_filter = set(str(p) for p in selected_priority)
        show_priorities_in_title = True
        show_priorities_in_filename = True
    elif selected_priority == "All":
        priorities_filter = None
        show_priorities_in_title = False
        show_priorities_in_filename = False
    elif selected_priority is None:
        priorities_filter = None
        show_priorities_in_title = False
        show_priorities_in_filename = False
    else:
        priorities_filter = set([str(selected_priority)])
        show_priorities_in_title = True
        show_priorities_in_filename = True

    for sheet_name in sheet_order:
        df = excel_data[sheet_name]
        df.columns = [col.strip().lower() for col in df.columns]
        if 'days spent to resolve' in df.columns:
            if priorities_filter is not None and 'priority' in df.columns:
                df = df[df['priority'].astype(str).isin(priorities_filter)]
            for days in df['days spent to resolve']:
                data.append({'Category': sheet_name, 'DaysToResolve': days})
        else:
            return False, None, 'No "days spent to resolve" column found.'

    if not data:
        if priorities_filter is not None:
            return False, None, "No valid data found in the Excel file for the selected priority."
        else:
            return False, None, "No valid data found in the Excel file."

    # Create DataFrame
    df = pd.DataFrame(data)

    # Generate box plot
    plt.figure(figsize=(8, 6))
    title = 'Days Spent to Resolve Defects by Category'
    if show_priorities_in_title and priorities_filter:
        title += f" (Priority: {', '.join(priorities_filter)})"
    sns.boxplot(x='Category', y='DaysToResolve', data=df, hue='Category', palette='Set2', order=sheet_order, legend=False)
    plt.title(title)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.ylabel('Days Spent to Resolve')
    plt.xlabel('Defect Category')
    plt.tight_layout()
    if show_priorities_in_filename and priorities_filter:
        output_path = f"{os.path.splitext(excel_file)[0]}_boxplot_{'_'.join(priorities_filter).lower()}.png"
    else:
        output_path = f"{os.path.splitext(excel_file)[0]}_boxplot.png"
    plt.savefig(output_path)
    plt.close()
    return True, output_path, None


def generate_priority_category_bar_plot(excel_file):
    """
    Generate a bar plot showing the relationship between defect priority and category.
    Each sheet is treated as a category. The function expects a categorized defect report Excel file.
    """
    # Read all sheets and concatenate into a single DataFrame with a 'Category' column
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    records = []
    for sheet_name, df in excel_data.items():
        if df.empty:
            continue
        df = df.copy()
        df.columns = [col.lower() for col in df.columns]
        if 'priority' in df.columns:
            for _, row in df.iterrows():
                records.append({'Category': sheet_name, 'Priority': row['priority']})
        else:
            raise ValueError("No 'priority' column found in the file.")
    plot_df = pd.DataFrame(records)
    # Set plot style
    sns.set_style("whitegrid")
    # Create count plot
    plt.figure(figsize=(10, 6))
    ax = sns.countplot(data=plot_df, x='Category', hue='Priority', palette='Set2')
    # Add title and labels
    plt.title('Defect Count by Category and Priority')
    plt.xlabel('Defect Category')
    plt.ylabel('Count')
    plt.legend(title='Priority')
    # Save plot
    plt.tight_layout()
    output_path = f"{os.path.splitext(excel_file)[0]}_priority_category_barplot.png"
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    return output_path


def generate_defect_type_category_bar_plot(excel_file):
    """
    Generate a bar plot showing the number of defects for defect types within each category.
    Each sheet is treated as a category. The function expects a categorized defect report Excel file.
    """
    # Read all sheets and concatenate into a single DataFrame with a 'Category' column
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    records = []
    for sheet_name, df in excel_data.items():
        if df.empty:
            continue
        df = df.copy()
        df.columns = [col.lower() for col in df.columns]
        if 'custom field (category)' in df.columns:
            for _, row in df.iterrows():
                records.append({'Category': sheet_name, 'DefectType': row['custom field (category)']})
        else:
            raise ValueError("No 'custom field (category)' column found in the file.")
    plot_df = pd.DataFrame(records)
    # Set plot style
    sns.set(style="whitegrid")
    # Create count plot
    plt.figure(figsize=(10, 6))
    ax = sns.countplot(data=plot_df, x='Category', hue='DefectType', palette='Set2')
    # Add title and labels
    plt.title('Defect Count by Category and Defect Type')
    plt.xlabel('Defect Category')
    plt.ylabel('Count')
    plt.legend(title='Defect Type')
    # Save plot
    plt.tight_layout()
    output_path = f"{os.path.splitext(excel_file)[0]}_defecttype_category_barplot.png"
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close()
    return output_path
