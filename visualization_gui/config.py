import os
# Modern Color Scheme
primary_color = "#3498db"       # Vibrant blue
primary_dark = "#2980b9"        # Darker blue for hover
secondary_color = "#f5f7fa"     # Light background (off-white)
accent_color = "#e74c3c"        # Red accent for important elements
success_color = "#2ecc71"       # Green for success states
error_color = "#e74c3c"         # Red for errors (same as accent for consistency)
text_color = "#2c3e50"          # Dark blue-gray for text
light_text = "#ecf0f1"          # Light text for dark backgrounds
highlight_color = "#d6eaf8"     # Light blue highlight
card_bg = "#ffffff"             # White for card backgrounds
border_color = "#d1d8e0"        # Light gray for borders

# Fonts
title_font = ("Segoe UI", 14, "bold") if os.name == 'nt' else ("Arial", 14, "bold")
header_font = ("Segoe UI", 12, "bold") if os.name == 'nt' else ("Arial", 12, "bold")
button_font = ("Segoe UI", 11) if os.name == 'nt' else ("Arial", 11)
small_button_font = ("Segoe UI", 10) if os.name == 'nt' else ("Arial", 10)
label_font = ("Segoe UI", 10) if os.name == 'nt' else ("Arial", 10)

# Dimensions and Padding
button_padding = (10, 5)
small_button_padding = (8, 4)
frame_padding = 15
section_padding = (8, 12)
dropdown_padding = (5, 2)

# Options
top_n_options = [("Top 10", 10), ("Top 20", 20), ("Top 40", 40), 
                 ("Top 60", 60), ("Top 80", 80), ("Top 100", 100)]