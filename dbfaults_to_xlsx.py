# dbfaults_to_xlsx.py v1.0 - Parse through Cisco's ACI concrete Fault DB and output to Xlsx file.
# Copyright (C) 2020  Scott Honey (scott.honey@thinkahead.com)
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#

# ACI db_faults from Cisco.com website, this will need a creation script to grab and convert the file.
import aci_fault_db

# This is the only python library we use, use pip3 to install it
import xlsxwriter


############################################################################################
# Global Namespace declarations
############################################################################################


class MO:
    name = ""
    Fault_Cons = {}

    def __init__(self, name):
        self.name = name
        self.Fault_Cons = {}  # probably pointless.
        return


MOs = {}


############################################################################################
# worksheet functions
############################################################################################

def create_worksheet_allfaults(workbook):
    # All the column #, column width, and the associated fcode field data.
    # It is possible to exclude any entry in the list, but ensure the column # is sequential and the commas matter.
    header_list = [
        [0, 15, 'Fault Code'],
        [1, 15, 'Fault Name'],
        [2, 15, 'Message'],
        [3, 15, 'Raised on MO'],
        [4, 15, 'Type'],
        [5, 15, 'Severity'],
        [6, 15, 'Cause'],
        [7, 15, 'Explanation'],
        [8, 15, 'Recommended Action'],
        [9, 15, 'Unqualified API Name'],
        [10, 15, 'Triggered By'],
        [11, 15, 'Applied MO DN Format']
    ]

    # Create our workbook and set the header row font to Bold.
    worksheet_faults = workbook.add_worksheet("All ACI Faults")
    header_format = workbook.add_format({'bold': True})

    # Write in the sheet header based on the header_list.
    current_row = 0
    for column, width, data in header_list:
        # header_list controls the header names for each column.
        # Set the column width and write the header text to the cell.
        worksheet_faults.set_column(column, column, width)
        worksheet_faults.write(current_row, column, data, header_format)

    current_row = 1  # Fill the rest of the sheet with data starting at row 1.

    # Iterate through the entire db_faults.
    for code in aci_fault_db.db_faults:
        # header_list controls which columns are created in the xlsx for this sheet.
        for column, width, data in header_list:
            # Write the column data if it's in the list.
            worksheet_faults.write(current_row, column, str(code[data]))

        # Next
        current_row += 1

    return


def create_worksheet_faultsbymo(workbook, worksheet_name=None, mo_list=None, sev_list=None, header_override=None):
    # Setup the default worksheet name.
    if worksheet_name is None:
        worksheet_name = "All ACI Faults by MO"

    # Default Header List: All the column #, column width, and the associated fcode field data.
    # It is possible to exclude any entry in the list, but ensure the column # is sequential and the commas matter.
    # Special Note: Excluding 'Raised on MO' in this report would make this report useless.
    if header_override is None:
        header_list = [
            [0, 33, 'Raised on MO'],
            [1, 7, 'Severity'],
            [2, 9, 'Fault Code'],
            [3, 58, 'Fault Name'],
            [4, 15, 'Message'],
            [5, 14, 'Type'],
            [6, 25, 'Cause'],
            [7, 255, 'Explanation'],
            [8, 15, 'Recommended Action'],
            [9, 15, 'Unqualified API Name'],
            [10, 15, 'Triggered By'],
            [11, 15, 'Applied MO DN Format']
        ]
    else:
        header_list = header_override

    # Default Severity List
    if sev_list is None:
        severity_list = ["critical", "major", "minor", "warning", "variable"]
    else:
        severity_list = sev_list

    # Create our workbook and set the header row font to Bold.
    worksheet = workbook.add_worksheet(worksheet_name)
    header_format = workbook.add_format({'bold': True})

    # Write in the sheet header based on the header_list.
    current_row = 0
    for column, width, data in header_list:
        # Set the column width and write the header text to the cell.
        worksheet.set_column(column, column, width)
        worksheet.write(current_row, column, data, header_format)

    current_row = 1  # Fill the rest of the sheet with data starting at row 1.

    if len(MOs.keys()) > 0:  # Check if we have MOs in our global dictionary.
        for mo in MOs.keys():  # Run through the list of dictionary keys
            # This is an optional flag.
            # If set during function call, only process MO class objects if the name is in the list.
            if mo_list is not None:
                if mo not in mo_list:
                    continue

            # Column 0 will be the Monitoring Object name in this sheet.
            worksheet.write(current_row, 0, MOs[mo].name)

            current_row += 1  # Skip to the next row after starting a new Monitoring Object entry.

            # This variable for tracking if we've found any fcodes matching our Severity List.
            # If not, we need to clean up
            found_fcode = False

            # Only process fcode severity levels in our list, inherit function default if not overwritten.
            for sev in severity_list:
                for fault in MOs[mo].Fault_Cons.keys():
                    # Check if the fault matches the current list entry.
                    if MOs[mo].Fault_Cons[fault]['Severity'] == sev:
                        # We've got a fault that needs to be added to the sheet, set found_fcode to True
                        found_fcode = True

                        # Only process column entries if it's in our header_list.
                        for column, width, data in header_list:
                            if column == 0:
                                continue  # Skip over Raised on MO column
                            else:
                                worksheet.write(current_row, column, str(MOs[mo].Fault_Cons[fault][data]))

                        # continue to next row and next fault
                        current_row += 1

            # if we didn't find any fcodes matching our sev_list for the entire MO, \
            # go back one row in the sheet and zero out column 0
            if not found_fcode:
                current_row -= 1
                worksheet.write(current_row, 0, "")
    return


############################################################################################
# Main()
############################################################################################

def main():
    print("Cisco ACI Fault Codes Imported: ", len(aci_fault_db.db_faults))

    # Fill the global MOs dictionary with Monitoring Objects from the db_faults.
    for fcode in aci_fault_db.db_faults:
        # Check if the MO is already in the dictionary. If it isn't,
        # create a new MO class object and add it to the dictionary.
        if fcode["Raised on MO"] not in MOs.keys():
            # This line looks like we're recursively assigning a variable to itself.
            MOs[fcode["Raised on MO"]] = MO(fcode["Raised on MO"])

        # Once we know we have a new MO class or
        MOs[fcode["Raised on MO"]].Fault_Cons[fcode["Fault Code"]] = fcode

    # Create the Excel file
    workbook = xlsxwriter.Workbook('aci-faults.xlsx')

    # EXAMPLE: Create a worksheet with all faults.
    # create_worksheet_allfaults(workbook)

    # EXAMPLE: Create a worksheet with all faults grouped by MO.
    # create_worksheet_faultsbymo(workbook)

    # EXAMPLE: Create a worksheet with exclude variable faults and grouped by MO.
    header_list = [
        [0, 33, 'Raised on MO'],
        [1, 7, 'Severity'],
        [2, 9, 'Fault Code'],
        [3, 58, 'Fault Name'],
        [4, 15, 'Message'],
        [5, 14, 'Type'],
        [6, 25, 'Cause'],
        [7, 255, 'Explanation']  # ,
        # [8, 15, 'Recommended Action'],
        # [9, 15, 'Unqualified API Name'],
        # [10, 15, 'Triggered By'],
        # [11, 15, 'Applied MO DN Format']
        ]

    create_worksheet_faultsbymo(workbook, worksheet_name="Fault Sev Assignment Policies",
                                sev_list=['critical', 'major', 'minor', 'warning'], header_override=header_list)

    # EXAMPLE: Create a worksheet off of a list of MOs. Helps to shrink the size of the spreadsheet
    # and offers better performance, at the cost of more spreadsheets....
    # Report_MOs=['eqpt:Storage','eqpt:Psu','vpc:If','ethpm:If']
    # Report_Severity=['critical','major','minor','warning']
    # create_worksheet_faultsbymo(workbook, worksheet_name="search_MO", mo_list=Report_MOs, sev_list=Report_Severity)

    # Close the workbook out.
    workbook.close()


if __name__ == "__main__":
    main()
