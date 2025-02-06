import pandas as pd
import os
import logging
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import tkinter as tk
from tkinter import filedialog
import dashboard as db
import time

log_folder = "Logs"
os.makedirs(log_folder, exist_ok=True)
# Set up logging
logging.basicConfig(
    level=logging.DEBUG, 
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('Logs/Comparison.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger()

# Declare global variables
file_entry_previous = None
file_entry_latest = None

calculated_amount = []
# calculated_totals = 0

def load_sheets(file_previous, file_latest):
    try:
        logger.info("Loading sheets from the provided Excel files.")
        sheet_pr_previous = pd.read_excel(file_previous, sheet_name="PR")
        sheet_pr_latest = pd.read_excel(file_latest, sheet_name="PR")
        sheet_vdpmv_previous = pd.read_excel(file_previous, sheet_name="SumVDPMVReport")
        sheet_vdpmv_latest = pd.read_excel(file_latest, sheet_name="SumVDPMVReport")
        sheet_partner_previous = pd.read_excel(file_previous, sheet_name="Div2PartnerList")
        sheet_partner_latest = pd.read_excel(file_latest, sheet_name="Div2PartnerList")
        logger.info("Sheets loaded successfully.")
        return sheet_pr_previous, sheet_pr_latest, sheet_vdpmv_previous, sheet_vdpmv_latest, sheet_partner_previous, sheet_partner_latest
    except Exception as e:
        logger.error(f"Error loading sheets: {e}")
        raise

def clean_currency(value):
    try:
        if isinstance(value, str):
            value = value.replace('$', '').replace(',', '').strip()
            return round(float(value), 2) if value else None
        return value
    except ValueError:
        logger.error(f"Error cleaning currency value: {value}")
        return None

def calculate_totals(hours_sheet, pr_sheet):
    try:
        calculated_totals = 0
        clients = hours_sheet["CLIENT"].unique()
        for client in clients: 
            # logger.info(f"Calculating totals for {client} clients.")

            client_header_row = pr_sheet[pr_sheet.iloc[:, 0].astype(str).str.contains(client, na=False, case=False)]
            # Get the index of the client header row
            client_header_index = client_header_row.index[0]
            # logger.info(f"Found client header for {client} at row {client_header_index}.")

            # Find the first empty row after the client header row to determine the endpoint
            empty_row_index = pr_sheet.iloc[client_header_index + 1:, 0].isna().idxmax()

            if pd.isna(empty_row_index):
                # If no empty row is found, use the last row of the DataFrame
                next_header_index = pr_sheet.shape[0]
            else:
                # The index of the first empty row marks the endpoint
                next_header_index = empty_row_index + client_header_index + 1

            # logger.info(f"Partner rows for client {client} are between rows {client_header_index + 1} and {next_header_index - 1}.")

            # Extract the partner rows between the client header and the first empty row
            partner_rows = pr_sheet.iloc[client_header_index + 1:next_header_index]

            # Get all unique partners for this client from the hours sheet
            total_partners = hours_sheet["PARTNER"].unique()
            # logger.info(f"Total partners for {client}: {len(total_partners)}")

            total_amount = 0

            # Calculate total amount for matching partners
            for partner in total_partners:
                # Find rows in partner_rows matching the partner
                partner_rows_matched = partner_rows[partner_rows.iloc[:, 0].astype(str).str.strip() == str(partner).strip()]

                if not partner_rows_matched.empty:
                    # Add the amount found in column 14 (assuming it's the amount column)
                    total_amount += partner_rows_matched.iloc[0, 14]  # Column 14 contains the amount
                    # logger.info(f"Amount for partner {partner}: {partner_rows_matched.iloc[0, 14]}")
            calculated_totals += total_amount
        logger.info(f"Total amount for client {client}: {total_amount}")
        # logger.info("Calculating totals for trips, hours, operators, and amounts.")
        logger.info(f"Total calculated mamount:{calculated_totals}")
        return {
            "TRIPS": hours_sheet["TRIPS"].sum(),
            "HOURS": hours_sheet["SERVICE HOURS OPERATED"].sum(),
            "OPERATORS": hours_sheet["OPERATOR NAME"].nunique(),
            "DAYS": hours_sheet["Date"].nunique(),
            "AMOUNT": calculated_totals
        }
    except Exception as e:
        logger.error(f"Error calculating totals: {e}")
        raise
    # try:
    #     total_amount = 0
    #     for totals in calculated_amount:
    #         print(totals)
    #         total_amount+=totals
        
    #     logger.info(f"Calculated Amounts: {total_amount}")
    #     logger.info("Calculating totals for trips, hours, operators, and amounts.")
    #     return {
    #         "TRIPS": hours_sheet["TRIPS"].sum(),
    #         "HOURS": hours_sheet["SERVICE HOURS OPERATED"].sum(),
    #         "OPERATORS": hours_sheet["OPERATOR NAME"].nunique(),
    #         "DAYS": hours_sheet["Date"].nunique(),
    #         "AMOUNT": total_amount
    #     }
    # except Exception as e:
    #     logger.error(f"Error calculating totals: {e}")
    #     raise

def calculate_client_totals(hours_sheet, pr_sheet, client):
    global calculated_amount
    try:
        # Search for the row index where the client header appears
        client_header_row = pr_sheet[pr_sheet.iloc[:, 0].astype(str).str.contains(client, na=False, case=False)]

        if client_header_row.empty:
            logger.warning(f"Client {client} not found in PR sheet. Returning 0 for amount.")
            return {
                "TRIPS": hours_sheet["TRIPS"].sum(),
                "HOURS": hours_sheet["SERVICE HOURS OPERATED"].sum(),
                "OPERATORS": hours_sheet["OPERATOR NAME"].nunique(),
                "DAYS": hours_sheet["Date"].nunique(),
                "AMOUNT": 0
            }

        # Get the index of the client header row
        client_header_index = client_header_row.index[0]
        logger.info(f"Found client header for {client} at row {client_header_index}.")

        # Find the first empty row after the client header row to determine the endpoint
        empty_row_index = pr_sheet.iloc[client_header_index + 1:, 0].isna().idxmax()

        if pd.isna(empty_row_index):
            # If no empty row is found, use the last row of the DataFrame
            next_header_index = pr_sheet.shape[0]
        else:
            # The index of the first empty row marks the endpoint
            next_header_index = empty_row_index + client_header_index + 1

        logger.info(f"Partner rows for client {client} are between rows {client_header_index + 1} and {next_header_index - 1}.")

        # Extract the partner rows between the client header and the first empty row
        partner_rows = pr_sheet.iloc[client_header_index + 1:next_header_index]

        # Get all unique partners for this client from the hours sheet
        total_partners = hours_sheet["PARTNER"].unique()
        logger.info(f"Total partners for {client}: {len(total_partners)}")

        total_amount = 0

        # Calculate total amount for matching partners
        for partner in total_partners:
            # Find rows in partner_rows matching the partner
            partner_rows_matched = partner_rows[partner_rows.iloc[:, 0].astype(str).str.strip() == str(partner).strip()]

            if not partner_rows_matched.empty:
                # Add the amount found in column 14 (assuming it's the amount column)
                total_amount += partner_rows_matched.iloc[0, 14]  # Column 14 contains the amount
                logger.info(f"Amount for partner {partner}: {partner_rows_matched.iloc[0, 14]}")
        calculated_amount.append(total_amount)
        logger.info(f"Total amount for client {client}: {total_amount}")

        # Return the calculated totals
        return {
            "TRIPS": hours_sheet["TRIPS"].sum(),
            "HOURS": hours_sheet["SERVICE HOURS OPERATED"].sum(),
            "OPERATORS": hours_sheet["OPERATOR NAME"].nunique(),
            "DAYS": hours_sheet["Date"].nunique(),
            "AMOUNT": total_amount
        }

    except Exception as e:
        logger.error(f"Error calculating client totals for {client}: {e}")
        raise

def compare_operators(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing operators between previous and latest sheets.")
        operators_previous = set(sheet_previous[["OPERATOR NAME", "PARTNER"]].dropna().itertuples(index=False, name=None))
        operators_latest = set(sheet_latest[["OPERATOR NAME", "PARTNER"]].dropna().itertuples(index=False, name=None))

        added = operators_latest - operators_previous
        removed = operators_previous - operators_latest

        added_list = [{"Operator Name": op, "Partner": partner} for op, partner in added]
        removed_list = [{"Operator Name": op, "Partner": partner} for op, partner in removed]

        logger.info("Operator comparison completed.")
        return {"Added": added_list, "Removed": removed_list}
    except Exception as e:
        logger.error(f"Error comparing operators: {e}")
        raise

def compare_dates(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing dates between previous and latest sheets.")

        # Extracting the Date column and dropping NaN values
        dates_previous = set(sheet_previous["Date"].dropna())
        dates_latest = set(sheet_latest["Date"].dropna())

        # Identifying added and removed dates
        added_dates = dates_latest - dates_previous
        removed_dates = dates_previous - dates_latest

        # Formatting the results as lists of dictionaries
        added_list = [{"Date": date} for date in added_dates]
        removed_list = [{"Date": date} for date in removed_dates]

        logger.info("Date comparison completed.")
        return {"Added": added_list, "Removed": removed_list}
    except Exception as e:
        logger.error(f"Error comparing dates: {e}")
        raise

def compare_trips_and_hours(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing trips and hours data between previous and latest sheets.")
        grouped_previous = sheet_previous.groupby("PARTNER")[["TRIPS", "SERVICE HOURS OPERATED"]].sum()
        grouped_latest = sheet_latest.groupby("PARTNER")[["TRIPS", "SERVICE HOURS OPERATED"]].sum()

        comparison = grouped_previous.join(grouped_latest, how="outer", lsuffix="_PREVIOUS", rsuffix="_LATEST").fillna(0)
        comparison["TRIPS_CHANGE"] = comparison["TRIPS_LATEST"] - comparison["TRIPS_PREVIOUS"]

        # Round hours values to two decimal places
        comparison["SERVICE HOURS OPERATED_PREVIOUS"] = comparison["SERVICE HOURS OPERATED_PREVIOUS"].round(2)
        comparison["SERVICE HOURS OPERATED_LATEST"] = comparison["SERVICE HOURS OPERATED_LATEST"].round(2)
        comparison["HOURS_CHANGE"] = (comparison["SERVICE HOURS OPERATED_LATEST"] - comparison["SERVICE HOURS OPERATED_PREVIOUS"]).round(2)

        trips_comparison = comparison[["TRIPS_PREVIOUS", "TRIPS_LATEST", "TRIPS_CHANGE"]].reset_index()
        trips_comparison.columns = ["PARTNER", "PREVIOUS", "LATEST", "CHANGE"]

        hours_comparison = comparison[["SERVICE HOURS OPERATED_PREVIOUS", "SERVICE HOURS OPERATED_LATEST", "HOURS_CHANGE"]].reset_index()
        hours_comparison.columns = ["PARTNER", "PREVIOUS", "LATEST", "CHANGE"]

        logger.info("Trips and hours comparison completed.")
        return trips_comparison, hours_comparison
    except Exception as e:
        logger.error(f"Error comparing trips and hours: {e}")
        raise

def compare_deductions(sheet_previous, sheet_latest):
    try:
        logger.info("Comparing Deductions between previous and latest sheets.")
        
        # Select the relevant columns directly
        previous_values = sheet_previous[["PARTNER", "LIFT LEASE TOTAL"]]
        latest_values = sheet_latest[["PARTNER", "LIFT LEASE TOTAL"]]

        # Merge both dataframes on "PARTNER" and handle missing values with 0
        comparison = previous_values.merge(latest_values, on="PARTNER", how="outer", suffixes=("_PREVIOUS", "_LATEST")).fillna(0)

        # Calculate the change in the "LIFT LEASE TOTAL"
        comparison["CHANGE"] = comparison["LIFT LEASE TOTAL_LATEST"] - comparison["LIFT LEASE TOTAL_PREVIOUS"]

        # Prepare the final dataframe for comparison
        deductions_comparison = comparison[["PARTNER", "LIFT LEASE TOTAL_PREVIOUS", "LIFT LEASE TOTAL_LATEST", "CHANGE"]].drop_duplicates(subset="PARTNER")

        # Rename columns
        deductions_comparison.columns = ["PARTNER", "PREVIOUS", "LATEST", "CHANGE"]

        logger.info("Deductions comparison completed.")
        return deductions_comparison

    except Exception as e:
        logger.error(f"Error comparing deductions: {e}")
        raise


def apply_formatting(sheet_name, wb):
    try:
        logger.info(f"Applying formatting to sheet: {sheet_name}.")
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                if ws[1][cell.column - 1].value.lower() == "change":
                    if isinstance(cell.value, (int, float)):
                        if cell.value > 0:
                            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            cell.font = Font(color="006100")
                        elif cell.value < 0:
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            cell.font = Font(color="9C0006")
                    elif isinstance(cell.value, str):
                        if cell.value.lower() in ["increased", "added"]:
                            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                            cell.font = Font(color="006100")
                        elif cell.value.lower() in ["decreased", "removed"]:
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                            cell.font = Font(color="9C0006")
        logger.info(f"Formatting applied successfully to sheet: {sheet_name}.")
    except Exception as e:
        logger.error(f"Error applying formatting to sheet {sheet_name}: {e}")
        raise


def save_comparison_results(output_folder, comparison_data, filename):
    try:
        logger.info(f"Saving comparison results to {filename}.")
        os.makedirs(output_folder, exist_ok=True)
        full_comparison_file = os.path.join(output_folder, filename)
        with pd.ExcelWriter(full_comparison_file, engine="openpyxl") as writer:
            for sheet_name, data in comparison_data.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        wb_full = load_workbook(full_comparison_file)
        for sheet in comparison_data.keys():
            apply_formatting(sheet, wb_full)
        wb_full.save(full_comparison_file)
        wb_full.close()

        logger.info(f"Comparison results saved successfully to {filename}.")
    except Exception as e:
        logger.error(f"Error saving comparison results: {e}")
        raise


def main(file_previous, file_latest):
    try: 
        file_entry_previous = file_previous
        file_entry_latest = file_latest

        print(f"{file_entry_previous} + {file_entry_latest}")
        logger.info("Starting main comparison process.")
        output_folder = "ComparedResults"
        os.makedirs(output_folder, exist_ok=True)

        sheet_pr_previous, sheet_pr_latest, sheet_hours_previous, sheet_hours_latest, sheet_lease_previous, sheet_lease_latest = load_sheets(file_previous, file_latest)
        # totals_previous = None
        # totals_latest = None
        # 2. Process the data for each client
        # Process data for each client
        unique_clients = sheet_hours_latest["CLIENT"].dropna().unique()

        for client in unique_clients:
            # Filter by client for both previous and latest sheets
            sheet_previous_client = sheet_hours_previous[sheet_hours_previous["CLIENT"] == client]
            sheet_latest_client = sheet_hours_latest[sheet_hours_latest["CLIENT"] == client]
            sheet_lease_previous_client = sheet_lease_previous[sheet_lease_previous["Type"] == client]
            sheet_lease_latest_client = sheet_lease_previous[sheet_lease_previous["Type"] == client]

            # Recalculate totals for client
            totals_previous_client = calculate_client_totals(sheet_previous_client, sheet_pr_previous, client)
            # totals_previous = calculate_totals(sheet_hours_previous, sheet_pr_previous)
            totals_latest_client = calculate_client_totals(sheet_latest_client, sheet_pr_latest, client)
            # totals_latest = calculate_totals(sheet_hours_latest, sheet_pr_latest)

            # Calculate differences and changes for the client
            differences_client = {
                "TRIPS": totals_latest_client["TRIPS"] - totals_previous_client["TRIPS"],
                "HOURS": totals_latest_client["HOURS"] - totals_previous_client["HOURS"],
                "OPERATORS": totals_latest_client["OPERATORS"] - totals_previous_client["OPERATORS"],
                "DAYS": totals_latest_client["DAYS"] - totals_previous_client["DAYS"],
                "AMOUNT": totals_latest_client["AMOUNT"] - totals_previous_client["AMOUNT"],
            }
            changes_client = {key: "Increased" if diff > 0 else "Decreased" if diff < 0 else "No Change"
                              for key, diff in differences_client.items()}

            # Create summary DataFrame for the client
            summary_table_client = {
                "Metric": ["TRIPS", "HOURS", "OPERATORS", "DAYS", "AMOUNT"],
                "Previous": [totals_previous_client[key] for key in totals_previous_client],
                "Latest": [totals_latest_client[key] for key in totals_latest_client],
                "Difference": [differences_client[key] for key in differences_client],
                "Change": [changes_client[key] for key in changes_client],
            }
            summary_df_client = pd.DataFrame(summary_table_client)

            # Compare operators for the client
            operator_changes_client = compare_operators(sheet_previous_client, sheet_latest_client)
            added_df_client = pd.DataFrame(operator_changes_client["Added"])
            removed_df_client = pd.DataFrame(operator_changes_client["Removed"])
            added_df_client["Change"] = "Added"
            removed_df_client["Change"] = "Removed"
            operator_changes_df_client = pd.concat([added_df_client, removed_df_client], ignore_index=True)

            # Compare operators
            # date_changes = compare_dates(sheet_hours_previous, sheet_hours_latest)
            # Dadded_df = pd.DataFrame(date_changes["Added"])
            # Dremoved_df = pd.DataFrame(date_changes["Removed"])
            # Dremoved_df["Change"] = "Removed"
            # Doperator_changes_df = pd.concat([Dadded_df, Dremoved_df], ignore_index=True)

            # Compare trips and hours for the client
            trips_comparison_df_client, hours_comparison_df_client = compare_trips_and_hours(sheet_previous_client, sheet_latest_client)
            lease_comparison_df = compare_deductions(sheet_lease_previous_client, sheet_lease_latest_client)

            # Save output for the client
            client_output_file = os.path.join(output_folder, f"{client}_Comparison.xlsx")
            with pd.ExcelWriter(client_output_file, engine="openpyxl") as writer:
                summary_df_client.to_excel(writer, sheet_name="Summary", index=False)
                operator_changes_df_client.to_excel(writer, sheet_name="OperatorChanges", index=False)
                trips_comparison_df_client.to_excel(writer, sheet_name="TripsComparison", index=False)
                hours_comparison_df_client.to_excel(writer, sheet_name="HoursComparison", index=False)
                lease_comparison_df.to_excel(writer, sheet_name="LeaseComparison", index=False)
                # Doperator_changes_df.to_excel(writer, sheet_name="DateComparison", index=False)

            # Apply formatting to the client's output
            wb_client = load_workbook(client_output_file)
            for sheet in ["Summary", "OperatorChanges", "TripsComparison", "HoursComparison","LeaseComparison"]:
                apply_formatting(sheet, wb_client)
            wb_client.save(client_output_file)
        wb_client.close()

        # 1. Process the data without any filtering (full comparison)
        # Recalculate totals for both previous and latest
        totals_previous = calculate_totals(sheet_hours_previous, sheet_pr_previous)
        totals_latest = calculate_totals(sheet_hours_latest, sheet_pr_latest)

        # Calculate differences and changes
        differences = {
            "TRIPS": totals_latest["TRIPS"] - totals_previous["TRIPS"],
            "HOURS": totals_latest["HOURS"] - totals_previous["HOURS"],
            "OPERATORS": totals_latest["OPERATORS"] - totals_previous["OPERATORS"],
            "DAYS": totals_latest["DAYS"] - totals_previous["DAYS"],
            "AMOUNT": totals_latest["AMOUNT"] - totals_previous["AMOUNT"],
        }
        changes = {key: "Increased" if diff > 0 else "Decreased" if diff < 0 else "No Change"
                    for key, diff in differences.items()}

        # Create summary DataFrame
        summary_table = {
            "Metric": ["TRIPS", "HOURS", "OPERATORS", "DAYS", "AMOUNT"],
            "Previous": [totals_previous[key] for key in totals_previous],
            "Latest": [totals_latest[key] for key in totals_latest],
            "Difference": [differences[key] for key in differences],
            "Change": [changes[key] for key in changes],
        }
        summary_df = pd.DataFrame(summary_table)

        # Compare operators
        operator_changes = compare_operators(sheet_hours_previous, sheet_hours_latest)
        added_df = pd.DataFrame(operator_changes["Added"])
        removed_df = pd.DataFrame(operator_changes["Removed"])
        added_df["Change"] = "Added"
        removed_df["Change"] = "Removed"
        operator_changes_df = pd.concat([added_df, removed_df], ignore_index=True)

        # Compare operators
        # date_changes = compare_dates(sheet_hours_previous, sheet_hours_latest)
        # Dadded_df = pd.DataFrame(date_changes["Added"])
        # Dremoved_df = pd.DataFrame(date_changes["Removed"])
        # Dadded_df["Change"] = "Added"
        # Dremoved_df["Change"] = "Removed"
        # Doperator_changes_df = pd.concat([Dadded_df, Dremoved_df], ignore_index=True)

        # Compare trips and hours
        trips_comparison_df, hours_comparison_df = compare_trips_and_hours(sheet_hours_previous, sheet_hours_latest)
        lease_comparison_df = compare_deductions(sheet_lease_previous, sheet_lease_latest)

        # Save the full comparison results
        full_comparison_file = os.path.join(output_folder, "Full_Comparison.xlsx")
        with pd.ExcelWriter(full_comparison_file, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
            operator_changes_df.to_excel(writer, sheet_name="OperatorChanges", index=False)
            trips_comparison_df.to_excel(writer, sheet_name="TripsComparison", index=False)
            hours_comparison_df.to_excel(writer, sheet_name="HoursComparison", index=False)
            lease_comparison_df.to_excel(writer, sheet_name="LeaseComparison", index=False)
            # Doperator_changes_df.to_excel(writer, sheet_name="DateComparison", index=False)

        # Apply formatting to the full comparison file
        wb_full = load_workbook(full_comparison_file)
        for sheet in ["Summary", "OperatorChanges", "TripsComparison", "HoursComparison", "LeaseComparison"]:
            apply_formatting(sheet, wb_full)
        wb_full.save(full_comparison_file)
        wb_full.close()

        logger.info(f"Main comparison process completed successfully. File saved to {full_comparison_file}.")
        time.sleep(2)
        db.main(file_previous, file_latest)
    except Exception as e:
        logger.error(f"Error in main comparison process: {e}")
        raise
    finally:
        wb_client.close()


def open_file_dialog(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xlsx")])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)


def create_gui():
    global file_entry_previous, file_entry_latest # Declare them as global variables

    # Set up the GUI window
    root = tk.Tk()
    root.title("Comparison Tool")

    # Create and place labels, entry boxes, and buttons
    tk.Label(root, text="Previous File:").grid(row=0, column=0, padx=10, pady=5)
    entry_previous = tk.Entry(root, width=50)
    entry_previous.grid(row=0, column=1, padx=10, pady=5)
    
    # tk.Button(root, text="Browse", command=lambda: open_file_dialog(entry_previous)).grid(row=0, column=2, padx=10, pady=5)

    tk.Label(root, text="Latest File:").grid(row=1, column=0, padx=10, pady=5)
    entry_latest = tk.Entry(root, width=50)
    entry_latest.grid(row=1, column=1, padx=10, pady=5)
    
    # tk.Button(root, text="Browse", command=lambda: open_file_dialog(entry_latest)).grid(row=1, column=2, padx=10, pady=5)

    # Button to trigger the comparison process
    # tk.Button(root, text="Compare", command=lambda: (main(entry_previous.get(), entry_latest.get()), root.destroy())).grid(row=2, column=1, pady=20)
    tk.Button(root, text="Compare", command=lambda: handle_comparison(entry_previous.get(), entry_latest.get(), root)).grid(row=2, column=1, pady=20)

    def handle_comparison(file_previous, file_latest, root):
        try:
            main(file_previous, file_latest)
        except Exception as e:
            print(f"An error occurred: {e}")
            # Check for the disconnection error and close the GUI if it happens
            if isinstance(e, OSError) and "The object invoked has disconnected" in str(e):
                print("Disconnected from Excel, closing GUI.")
                root.quit()  # This will close the Tkinter window
        finally:
            root.destroy()  # Close the window in all cases
    # Start the GUI loop
    root.mainloop()


if __name__ == "__main__":
    create_gui()

