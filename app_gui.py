import pandas as pd
import tkinter as tk
import customtkinter
import os
import re

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    # Application window width and height variables
    WIDTH = 1000
    HEIGHT = 550

    # Column and row indices
    COLUMN_INDEX = 1
    ROW_INDEX = 0

    # Keep track of right_container rows
    right_container_rows = 0

    # Application name
    APP_NAME = "Compatibility Matrix Master"

    # Base file path
    PATH = os.path.dirname(os.path.realpath(__file__))

    # Store Excel file path - default path is current project path
    file_path = PATH + "/Compatibility_Matrix_LST-1592.xlsx"

    # List of columns
    software_device_compatibility_labels_list = ['SafetyNet', 'MICT', 'Sketch', 'Trace', 'Tir-1',
                                                 'Radius-7', 'Radius-7 Wifi', 'Radius T', 'Centroid', 'VSM', 'SedLine',
                                                 'MOA', 'External SatShare', 'Iris - DMS', 'Iris Gateway']

    # Eagle version list
    eagle_versions_list = ['V1002 – V1303', 'V1412', 'V1434', 'V1437',
                           'V1448', 'V1461', 'V1470', 'V1514', 'V1518',
                           'V1536', 'V1544',
                           'V1569', 'V1578', 'V1593', 'V1615', 'V1616',
                           'V1617', 'V1621',
                           'V1624', 'V1625', 'V1631', 'V1728', 'V1748',
                           'V1750', 'V1767',
                           'V1773', 'V1798', 'V1814', 'V1818', 'V1819',
                           'V1820', 'V1945',
                           'V1973', 'V1983', 'V1988', 'V1992', 'V2061',
                           'V2096', 'V2101',
                           'V2103', 'V2104', 'V2105', 'V2106', 'V2116',
                           'V2118', 'V2120[3]',
                           'V2122', 'V2124', 'V2131', 'V2135', 'V2140',
                           'V2142', 'V2145',
                           'V2146']

    # Dictionary that contains compatible exceptions.
    compatibility_exceptions_dictionary = {'[1]': 'Requires minimum SafetyNet V4400',
                                           '[2]': 'Requires minimum SafetyNet V5008',
                                           '[3]': 'Requires Sedline V1203 to support \nall features',
                                           '[4]': 'Requires minimum MICT V1049',
                                           '[5]': 'Requires minimum Radius-7 IB V1012 \nfor all features supported',
                                           '[6]': 'Minimum Eagle version to support Falcon-pro.',
                                           '[7]': 'Requires minimum MICT V1109',
                                           '[8]': 'Requires minimum Radius-7 V1020',
                                           '[9]': 'Requires minimum Trace V2026',
                                           '[10]': 'Requires minimum IB-Pro V205X, \nrequires minimum Radius-7 BB V2015 \nto support all features',
                                           '[11]': 'Requires minimum SedLine V2320 for \nall Eagle enhancements to be available',
                                           '[12]': 'Added support for Safety Net V5027-5085',
                                           '[13]': 'Minimum eagle version to support PSN V5647',
                                           '[14]': 'Requires minimum IB-Pro V206x and \nminimum IB V103x with minimum BBV202x \nto support all features',
                                           '[15]': 'Minimum Eagle version to support PSN V5672',
                                           '[16]': 'Requires minimum Trace V2.0.2.8',
                                           '[17]': 'Minimum Eagle version to support Iris Gateway V1613',
                                           '[18]': 'Minimum Eagle version to support PSN V5675',
                                           '[19]': 'Requires minimum Trace V3025',
                                           '[20]': 'Minimum Eagle version to support VSM V1020',
                                           '[21]': 'V2120 is built from V2106',
                                           '[22]': 'Requires minimum MICT V1248 to support all features',
                                           '[23]': 'Requires minimum MICT V1252 to support all features',
                                           '[24]': 'Requires minimum MICT V1253 to support all features',
                                           '[25]': 'Requires minimum Trace V3025'}

    # Initialization method
    def __init__(self):
        # Initialize everything upon app creation contained in this section
        super().__init__()

        # Application title
        self.title(App.APP_NAME)

        # Set app windows width and height
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")

        # call .on_closing() when app gets closed
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # ============ Container(s)/Frame(s) Start ============

        '''

             - Configure grid layout (2x1) i.e. 2 columns and 1 row (col x row).
             - Note: weight parameter determines how wide the column will
                     occupy, which is relative to the columns.

        '''

        # The columnconfigure() method configures the column index of the grid
        # since we will have 2 columns the current index value is 1.  The columns
        # start from 0 - so 1 one is actually 2 columns and 0 is one.
        self.grid_columnconfigure(App.COLUMN_INDEX, weight=1)

        # The rowconfigure() method configures the row index of the grid
        # we will have only one row so the current index value is 0.
        self.grid_rowconfigure(App.ROW_INDEX, weight=1)

        # Configure left container/frame
        self.left_container_frame = customtkinter.CTkFrame(master=self, width=0)
        self.left_container_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nwsw")

        # Configure center container/frame
        self.center_container_frame = customtkinter.CTkFrame(master=self, width=600, corner_radius=15)
        self.center_container_frame.grid(row=0, column=1, padx=0, pady=20, sticky="ns")

        # Configure right container/frame
        self.right_container_frame = customtkinter.CTkFrame(master=self, width=0, corner_radius=15)
        self.right_container_frame.grid(row=0, column=2, padx=0, pady=0, sticky="nese")

        # ============ Container(s)/Frame(s) End ============

        # ============ Labels Start ============

        # Place welcome banner label in top center of center container
        self.banner_label = customtkinter.CTkLabel(master=self.center_container_frame,
                                                   text=self.APP_NAME,
                                                   text_font=("Roboto Large", -25))  # font name and size in px

        # Set the welcome banner labels grid configuration
        self.banner_label.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky='n')

        # Eagle Version Label
        self.eagle_version_label = customtkinter.CTkLabel(master=self.center_container_frame,
                                                          text="Select Eagle Version From \nThe Drop Down Menu",
                                                          height=200,
                                                          fg_color=("white", "gray38"),
                                                          text_font=('Verdana', 12))
        # Set the Eagle Version Label grid configuration
        self.eagle_version_label.grid(row=1, column=0, columnspan=2, pady=10, padx=5, sticky='n')

        self.radio_var = tk.IntVar()
        self.radio_var.set(0)

        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.center_container_frame,
                                                           text='Reverse',
                                                           command=self.reverse_eagle_version_list,
                                                           variable=self.radio_var,
                                                           value=0)
        self.radio_button_1.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="n")

        # ============ Labels End ============

        # ============ Option Menu Start ============

        # Eagle option menu
        self.eagle_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame,
                                                                     values=self.eagle_versions_list,
                                                                     command=self.option_menu_callback)
        self.eagle_version_option_menu.grid(row=2, column=0, columnspan=2, pady=10, padx=5)

        # ============ Combo Box(s) End ============

    # ============ Methods Start ============

    '''
        Method: on_closing
        Purpose: Destroy the entire application window when exit code = 0.
    '''

    def on_closing(self, event=0):
        return self.destroy()  # Destroy the entire application i.e. close out

    '''
         Method: option_menu_callback
         Purpose: Once a choice (Eagle version) is selected from the drop down menu,
                  the left/right container(s) will be destroyed i.e. all current values will
                  be erased on screen and updated accordingly.
     '''

    def option_menu_callback(self, choice):

        # Erase current values from left container on screen.
        self.left_container_frame.destroy()

        # Configure left container/frame
        self.left_container_frame = customtkinter.CTkFrame(master=self, width=0)
        self.left_container_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nwsw")

        # Erase current values from right contained on screen
        self.right_container_frame.destroy()

        # Configure right container/frame
        self.right_container_frame = customtkinter.CTkFrame(master=self, width=0, corner_radius=15)
        self.right_container_frame.grid(row=0, column=2, padx=0, pady=0, sticky="nese")

        # Index will access each column(s) 'yes' or 'no' value(s)
        index = 0

        # Read the Excel file into dataframe convert matrix to dictionary and return results
        matrix_dictionary = self.process_excel_file()

        # Pass the user's choice as 'key' to matrix/dictionary
        compatibility_list = matrix_dictionary[choice]

        # Find exceptions if they exist for the Eagle version
        eagle_version_exception = self.find_exceptions_regex(choice)

        # If an eagle version exception exists i.e. the result is not empty
        if eagle_version_exception != '':
            # Display the exception in the right container on screen
            self.create_label_right_container(choice, eagle_version_exception)

        # Iterate over each device in the compatibility list
        for device_yes_no_value in compatibility_list:

            # Get the device/service name from software/device compatibility list,
            # the 'yes' or 'no' value for the related device/service, and index to pass as row value.
            self.create_label_left_container(self.software_device_compatibility_labels_list[index], device_yes_no_value,
                                             index)

            # Find exceptions if they exist for the device/service version
            device_version_exception = self.find_exceptions_regex(device_yes_no_value)

            # If a device/service version exception exists i.e. the result is not empty
            if device_version_exception != '':
                self.create_label_right_container(self.software_device_compatibility_labels_list[index],
                                                  device_version_exception)

            index += 1  # Increase index count by one

        # Set right container rows back to 0
        self.right_container_rows = 0

        return choice  # return users Eagle version choice

    '''
        Method: open_excel_file
        Purpose: Open the Excel file, find the second table (header=16) located in the Excel sheet, 
                 and load columns B through Q (usecols='B:Q'), load rows (nrows=54) into a pandas dataframe.
    '''

    def open_excel_file(self):
        return pd.read_excel(self.file_path, header=16, usecols='B:Q', nrows=54)

    '''
        Method: get_columns
        Purpose: Simple helper method that returns the columns in a dataframe.
    '''

    def get_columns(self, df):
        return df.columns

    '''
        Method: convert_dataframe_to_numpy_array
        Purpose: Convert pandas dataframe to numpy array.
    '''

    def convert_dataframe_to_numpy_array(self, df):
        return df.to_numpy()

    '''
          Method: convert_excel_table_to_numpy_matrix
          Purpose: Convert Excel table to matrix/dictionary.
    '''

    def convert_excel_table_to_numpy_matrix(self, array_np):
        # Dictionary matrix
        matrix_dictionary = {}

        # Extract the eagle version as the key and compatibility matrix as the values
        for ls in array_np:
            # Grab eagle version as key
            key = ls[0]
            # Grab the rest of the list and use it as values
            values = ls[1:]
            # Set the key value pair
            matrix_dictionary[key] = values

        return matrix_dictionary  # Return the numpy matrix to caller

    '''
        Method: process_excel_file
        Purpose: Read the Excel file into a pandas dataframe
                 then convert the dataframe to numpy array.
    '''

    def process_excel_file(self):
        # This dataframe contains the second table starting with column 'Eagle Version'
        data_frame = self.open_excel_file()

        # Convert pandas dataframe to Numpy array
        array_np = self.convert_dataframe_to_numpy_array(data_frame)

        return self.convert_excel_table_to_numpy_matrix(array_np)  # Return numpy array

    '''
         Method: create_label_left_container
         Purpose: Create a label(s) that will be displayed in the left
                  container on screen.
    '''

    def create_label_left_container(self, compatibility_result, device, row):
        # Compatibility Label
        compatible_label = customtkinter.CTkLabel(master=self.left_container_frame,
                                                  text=compatibility_result + ": " + device,
                                                  text_font=('Verdana', 11),
                                                  fg_color=("white", "gray38"),
                                                  justify=tk.CENTER)

        # Set the Compatibility Version Label grid configuration
        compatible_label.grid(row=row, column=0, pady=4, padx=2)

    '''
         Method: create_label_right_container
         Purpose: Create a label(s) that will be displayed in the right
                  container on screen.
    '''

    def create_label_right_container(self, version, message):
        # Exceptions Label
        compatible_label = customtkinter.CTkLabel(master=self.right_container_frame,
                                                  text=version + ": " + message,
                                                  text_font=('Verdana', 11),
                                                  fg_color=("white", "gray38"),
                                                  justify=tk.CENTER)

        # Set the Exceptions Label grid configuration
        compatible_label.grid(row=self.right_container_rows, column=0, pady=10, padx=2)

        # Increment the number of rows my one
        self.right_container_rows += 1

    '''
         Method: find_exceptions_regex
         Purpose: to find pattern which contains one or two digits 
                  enclosed in brackets example [2] or [10]
    '''

    def find_exceptions_regex(self, string):
        # Find pattern which contains one or two digits enclosed in brackets example [2] or [10]
        pattern = re.compile(r'\[(\d){1,2}\]')

        # If the pattern is found in the string
        if pattern.search(string):
            # Grab the entire pattern example [2] or [10]
            result = pattern.search(string).group(0)
            # Return the results
            return self.compatibility_exceptions_dictionary[result]
        else:
            return ''  # Return an empty string

    '''
         Method: reverse_eagle_version_list
         Purpose: To allow the user to display oldest Eagle version
                  first or newest by reversing the Eagle versions list.
    '''

    def reverse_eagle_version_list(self):
        # Erase the current options menu versions
        self.eagle_version_option_menu.destroy()
        # Reverse the versions list
        self.eagle_versions_list.reverse()
        # Display the updated Eagle option menu
        self.eagle_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame,
                                                                     values=self.eagle_versions_list,
                                                                     command=self.option_menu_callback)
        self.eagle_version_option_menu.grid(row=2, column=0, columnspan=2, pady=10, padx=5)

    # ============ Methods Start ============
