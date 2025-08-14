# Defect Report Analysis Tool

This tool provides a graphical user interface (GUI) for analyzing, visualizing, and categorizing software defect reports. It supports advanced NLP, machine learning, and data visualization features for Excel defect reports from Jira.

## Features
- Filter defect reports by multiple columns (priority, category, created/resolved date)
- Add resolution period columns to reports
- Build and train a defect categorization model
- Categorize defect reports using a trained model
- Generate pie charts, box plots, bar plots, word count tables, word clouds, and bubble charts
- Japanese and English text processing
- Multi-sheet Excel support

## Installation
1. Clone or download 'visualization_gui' folder.
2. Open the folder with a code editor.
3. Install all required Python packages:

```powershell
pip install -r requirements.txt
```

## Usage
Run the main application:

```powershell
python main.py
```

Follow the GUI prompts to select Excel files and generate visualizations or models.

(Optional) For EXE creation, install PyInstaller:

```powershell
pip install pyinstaller
```
Then double click on the batch file to create an EXE file.

## Requirements
- Python 3.12.7 (tested)
- All packages listed in `requirements.txt`
- Windows OS recommended (tested)
- Japanese font file `src/ipaexg.ttf` (bundled)

## Main Functions
- **Filter Defect Reports**: Filter by priority, category, created/resolved date
- **Add Resolution Period Column**: Calculate days spent to resolve
- **Build Model for Categorization**: Train ML model on defect summaries
- **Categorize Defect Reports**: Apply trained model to new data
- **Generate Pie/Box/Bar Plots**: Visualize defect categories, priorities, issue types
- **Generate Word Count Table/Word Cloud/Bubble Chart**: Text analysis and visualization

## File Structure
- `main.py`: Entry point for the GUI
- `gui_app.py`: Main GUI logic
- `plots.py`: Visualization functions
- `summary_classifier.py`: ML-based categorization
- `build_model.py`: Model training
- `filter_defect_reports.py`: Filtering logic
- `resolution_column.py`: Add resolution period
- `word_count_util.py`, `text_preprocessing.py`: Text analysis
- `requirements.txt`: All required Python packages
- `src/ipaexg.ttf`: Japanese font
