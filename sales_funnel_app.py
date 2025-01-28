import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Web App Title
st.title("Quarterly KPI Visualizer")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xlsm"])

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

        # Loop through each row (quarter)
        for index, row in df.iterrows():
            # Extract data for the current quarter
            quarter = f"Q{index + 1} 2025"  # Title based on row index
            stages = [col.split()[0] for col in df.columns if 'ist' in col.lower()]  # Extract stage names
            actual_values = [row[f"{stage} (ist)"] for stage in stages]
            target_values = [row[f"{stage} (soll)"] for stage in stages]

            # Plot the KPI funnel for the current quarter
            fig, ax = plt.subplots(figsize=(12, 8))
            max_value = max(max(actual_values), max(target_values))  # Scale based on max value
            total_stages = len(stages)
            colors = ['red', 'yellow', 'purple', 'blue']

            prev_width_actual_top = actual_values[0] / max_value / 2
            prev_width_target_top = target_values[0] / max_value / 2

            for i, (stage, actual, target) in enumerate(zip(stages, actual_values, target_values)):
                y_top = (total_stages - i) / total_stages
                y_bottom = (total_stages - i - 1) / total_stages

                width_actual_top = prev_width_actual_top
                width_actual_bottom = max((actual / max_value / 2) * 0.8, 0.02)  # Ensure non-zero width

                width_target_top = prev_width_target_top
                width_target_bottom = max((target / max_value / 2) * 0.8, 0.02)  # Ensure non-zero width

                # Draw the "ist" areas
                ax.fill(
                    [0.5 - width_actual_top, 0.5 + width_actual_top, 0.5 + width_actual_bottom, 0.5 - width_actual_bottom],
                    [y_top, y_top, y_bottom, y_bottom],
                    color=colors[i % len(colors)],
                )

                # Draw the "soll" dashed lines
                ax.plot(
                    [0.5 - width_target_top, 0.5 + width_target_top, 0.5 + width_target_bottom, 0.5 - width_target_bottom, 0.5 - width_target_top],
                    [y_top, y_top, y_bottom, y_bottom, y_top],
                    linestyle='dashed', color='black'
                )

                # Add headline name in the filled shape
                y_center = (y_top + y_bottom) / 2  # Vertical center of the shape
                text_color = "black" if colors[i % len(colors)] == 'yellow' else "white"  # Black for yellow, white otherwise
                ax.text(
                    0.5, y_center,
                    stage,  # Use the stage name
                    va="center",
                    ha="center",
                    fontsize=12,
                    color=text_color,
                    fontweight="bold"
                )

                # Position `ist` measurement styles (on top of the filled shape)
                y_measurement_ist = y_top + 0.02  # Slightly above the top of the shape
                ax.plot(
                    [0.5 - width_actual_top, 0.5 + width_actual_top],
                    [y_measurement_ist, y_measurement_ist],
                    color="black",
                    linestyle="-",
                    linewidth=1
                )
                ax.annotate(
                    "", xy=(0.5 - width_actual_top, y_measurement_ist), xytext=(0.5 - width_actual_top - 0.02, y_measurement_ist),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1)
                )
                ax.annotate(
                    "", xy=(0.5 + width_actual_top, y_measurement_ist), xytext=(0.5 + width_actual_top + 0.02, y_measurement_ist),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1)
                )
                ax.text(
                    0.5, y_measurement_ist + 0.02,
                    f"Ist: {actual:.2f}",
                    va="bottom",
                    ha="center",
                    fontsize=10,
                    color="black"
                )

                # Position `soll` measurement styles (below the dashed line)
                y_measurement_soll = y_top - 0.02  # Slightly below the dashed line
                ax.plot(
                    [0.5 - width_target_top, 0.5 + width_target_top],
                    [y_measurement_soll, y_measurement_soll],
                    color="black",
                    linestyle="-",
                    linewidth=1
                )
                ax.annotate(
                    "", xy=(0.5 - width_target_top, y_measurement_soll), xytext=(0.5 - width_target_top - 0.02, y_measurement_soll),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1)
                )
                ax.annotate(
                    "", xy=(0.5 + width_target_top, y_measurement_soll), xytext=(0.5 + width_target_top + 0.02, y_measurement_soll),
                    arrowprops=dict(arrowstyle="->", color="black", lw=1)
                )
                ax.text(
                    0.5, y_measurement_soll - 0.02,
                    f"Soll: {target:.2f}",
                    va="top",
                    ha="center",
                    fontsize=10,
                    color="black"
                )

                prev_width_actual_top = width_actual_bottom
                prev_width_target_top = width_target_bottom

            # Set title and clean the chart
            ax.set_title(f"KPI - {quarter}", fontsize=16, loc='center', pad=20)
            ax.set_xlim(0, 1.4)
            ax.set_ylim(0, 1.3)
            ax.axis("off")

            # Display the chart
            st.pyplot(fig)

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
