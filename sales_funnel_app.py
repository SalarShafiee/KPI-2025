import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Streamlit App Title
st.set_page_config(page_title="KPI Visualizer", layout="wide")
st.title("📊 Vierteljährlicher KPI-Visualisierer")

# Upload Excel Button (Always Visible with Excel Icon)
st.markdown("### 📂 Laden Sie Ihre Excel-Datei hoch")
uploaded_file = st.file_uploader("Wählen Sie eine Excel-Datei", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # Read the Excel file
        excel_file = pd.ExcelFile(uploaded_file)

        # Check for correct sheet names
        if "Sheet2" in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name="Sheet2")
        elif "Tabelle (2)" in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name="Tabelle (2)")
        else:
            st.error("The file must contain a sheet named 'Sheet2' or 'Tabelle (2)'.")
            st.stop()

        # Make the table editable
        df_editable = st.data_editor(df, num_rows="dynamic")

        # Loop through each row (quarter)
        for index, row in df_editable.iterrows():
            quarter = f"Q{index + 1} 2025"
            stages = [col.split()[0] for col in df_editable.columns if 'ist' in col.lower()]
            actual_values = [row[f"{stage} (ist)"] for stage in stages]
            target_values = [row[f"{stage} (soll)"] for stage in stages]

            # Plot the KPI funnel for the current quarter
            fig, ax = plt.subplots(figsize=(12, 8))
            max_value = max(max(actual_values), max(target_values))
            total_stages = len(stages)
            colors = ['red', 'yellow', 'purple', 'blue']

            scaling_factor = 0.8  # Apply consistent scaling

            # Initialize the bottom widths of the first (topmost) shape
            width_actual_bottom = (actual_values[-1] / max_value / 2) * scaling_factor
            width_target_bottom = (target_values[-1] / max_value / 2) * scaling_factor

            for i in reversed(range(total_stages)):
                stage = stages[i]
                actual = actual_values[i]
                target = target_values[i]

                # Calculate the top widths of the current shape
                width_actual_top = (actual / max_value / 2) * scaling_factor
                width_target_top = (target / max_value / 2) * scaling_factor

                # Set the bottom widths of the current shape to match the top widths of the next shape (if applicable)
                if i < total_stages - 1:
                    width_actual_bottom = (actual_values[i + 1] / max_value / 2) * scaling_factor
                    width_target_bottom = (target_values[i + 1] / max_value / 2) * scaling_factor
                else:
                    # For the lowest shape, set the bottom widths to 0 (point)
                    width_actual_bottom = 0
                    width_target_bottom = 0

                y_top = (total_stages - i) / total_stages
                y_bottom = (total_stages - i - 1) / total_stages

                
                # Draw the "Ist" areas (trapezoids)
                ax.fill(
                    [0.5 - width_actual_top, 0.5 + width_actual_top, 0.5 + width_actual_bottom, 0.5 - width_actual_bottom],
                    [y_top, y_top, y_bottom, y_bottom],
                    color=colors[i % len(colors)],
                )

                # Draw the "Soll" dashed lines (trapezoids)
                ax.plot(
                    [0.5 - width_target_top, 0.5 + width_target_top, 0.5 + width_target_bottom, 0.5 - width_target_bottom, 0.5 - width_target_top],
                    [y_top, y_top, y_bottom, y_bottom, y_top],
                    linestyle='dashed', color='black'
                )

                # Add stage name inside the shape
                y_center = (y_top + y_bottom) / 2
                text_color = "black" if colors[i % len(colors)] == 'yellow' else "white"
                ax.text(0.5, y_center, stage, va="center", ha="center", fontsize=12, color=text_color, fontweight="bold")

                # Display "Ist" value annotations
                y_measurement_ist = y_top + 0.02
                ax.plot([0.5 - width_actual_top, 0.5 + width_actual_top], [y_measurement_ist, y_measurement_ist], color="black", linestyle="-", linewidth=1)
                ax.annotate("", xy=(0.5 - width_actual_top, y_measurement_ist), xytext=(0.5 - width_actual_top - 0.02, y_measurement_ist), arrowprops=dict(arrowstyle="->", color="black", lw=1))
                ax.annotate("", xy=(0.5 + width_actual_top, y_measurement_ist), xytext=(0.5 + width_actual_top + 0.02, y_measurement_ist), arrowprops=dict(arrowstyle="->", color="black", lw=1))
                ax.text(0.5, y_measurement_ist + 0.02, f"Ist: {actual:.2f}", va="bottom", ha="center", fontsize=10, color="black")

                # Display "Soll" value annotations
                y_measurement_soll = y_top - 0.02
                ax.plot([0.5 - width_target_top, 0.5 + width_target_top], [y_measurement_soll, y_measurement_soll], color="black", linestyle="-", linewidth=1)
                ax.annotate("", xy=(0.5 - width_target_top, y_measurement_soll), xytext=(0.5 - width_target_top - 0.02, y_measurement_soll), arrowprops=dict(arrowstyle="->", color="black", lw=1))
                ax.annotate("", xy=(0.5 + width_target_top, y_measurement_soll), xytext=(0.5 + width_target_top + 0.02, y_measurement_soll), arrowprops=dict(arrowstyle="->", color="black", lw=1))
                ax.text(0.5, y_measurement_soll - 0.02, f"Soll: {target:.2f}", va="top", ha="center", fontsize=10, color="black")

            # Set chart title and remove axes
            ax.set_title(f"KPI - {quarter}", fontsize=16, loc='center', pad=20)
            ax.set_xlim(0, 1.4)
            ax.set_ylim(0, 1.3)
            ax.axis("off")

            # Display the chart
            st.pyplot(fig)

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")