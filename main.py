# --------------------------------------------
# Python Program: Add Names and Numbers to Excel
# Made by Kasra Tookallo with an assistance of AI (problem_solving section).
# --------------------------------------------
# Description:
# This program repeatedly asks the user to enter a name and a number.
# Each entry is added to a list and displayed before saving.
# When the user types "save", all collected data is written to an Excel file.
# --------------------------------------------

# Importing required library
import openpyxl

# Initialize an empty list to store [name, number] pairs
data_list = []

print("=== Name and Number Collector ===")
print("Type 'save' at any time to export data to Excel and exit.\n")

# Infinite loop for continuous input
while True:
    # Get name input
    name = input("Enter a name (or type 'save' to finish): ").strip()

    # Exit condition
    if name.lower() == "save":
        break

    # Validate name input
    if not name:
        print("⚠️ Name cannot be empty. Try again.\n")
        continue

    # Get number input
    try:
        number = int(input("Enter a number for this name: "))
    except ValueError:
        print("⚠️ Invalid number. Please enter a valid integer.\n")
        continue

    # Append to list
    data_list.append([name, number])

    # Display current list
    print("\n✅ Current List:")
    for i, (n, num) in enumerate(data_list, start=1):
        print(f"{i}. {n} - {num}")
    print("\n-----------------------------\n")

# If list is empty, warn the user
if not data_list:
    print("⚠️ No data to save. Exiting without saving.")
    exit()

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Bahar Complex Members_List"

# Create headers
sheet["A1"] = "Name"
sheet["B1"] = "Payment"

# Write data to Excel
for row_index, (name, number) in enumerate(data_list, start=2):
    sheet[f"A{row_index}"] = name
    sheet[f"B{row_index}"] = number

# Save Excel file
excel_filename = "Bahar Complex Members_List.xlsx"
workbook.save(excel_filename)

print(f"\n💾 Data successfully saved to '{excel_filename}'.")
print("✅ Program finished.")




