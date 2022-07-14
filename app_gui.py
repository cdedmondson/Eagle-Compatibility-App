import pandas as pd
import tkinter as tk
from tkinter import Text
import customtkinter
import os
import re

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    # Application window width and height variables
    WIDTH = 600
    HEIGHT = 720

    # Column and row indices
    COLUMN_INDEX = 1
    ROW_INDEX = 0

    # Application name
    APP_NAME = "Eagle Compatibility Master"

    # Base file path
    PATH = os.path.dirname(os.path.realpath(__file__))

    # Store Excel file path - default path is current project path
    file_path = PATH + "/Compatibility_Matrix_LST-1592.xlsx"

    # Eagle version menu option will start off with 'V1002 – V1303'
    global_eagle_version_choice = 'V1002 – V1303'

    # Device/software version menu option will start off with 'SafetyNet'
    global_software_device_choice = 'SafetyNet'

    # Matrix dictionary
    matrix_dictionary = {}

    # List of columns
    software_device_list = ['SafetyNet', 'MICT', 'Sketch', 'Trace', 'Tir-1',
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
                                           '[3]': 'Requires Sedline V1203 to support all features',
                                           '[4]': 'Requires minimum MICT V1049',
                                           '[5]': 'Requires minimum Radius-7 IB V1012 for all features supported',
                                           '[6]': 'Minimum Eagle version to support Falcon-pro.',
                                           '[7]': 'Requires minimum MICT V1109',
                                           '[8]': 'Requires minimum Radius-7 V1020',
                                           '[9]': 'Requires minimum Trace V2026',
                                           '[10]': 'Requires minimum IB-Pro V205X, requires minimum Radius-7 BB V2015 to support all features',
                                           '[11]': 'Requires minimum SedLine V2320 for all Eagle enhancements to be available',
                                           '[12]': 'Added support for Safety Net V5027-5085',
                                           '[13]': 'Minimum eagle version to support PSN V5647',
                                           '[14]': 'Requires minimum IB-Pro V206x and minimum IB V103x with minimum BBV202x to support all features',
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
                                           '[25]': 'Requires minimum Trace V3025',
                                           '[26]': 'Requires minimum Centroid V2101'}

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
        # start from 0 - so a column value of 1, represents 2 columns total.
        self.grid_columnconfigure(App.COLUMN_INDEX, weight=1)

        # The rowconfigure() method configures the row index of the grid
        # we will have only one row, so the current index value is 0.
        self.grid_rowconfigure(App.ROW_INDEX, weight=1)

        # Configure left container/frame
        self.left_container_frame = customtkinter.CTkFrame(master=self, width=0)
        self.left_container_frame.grid(row=0, column=0, padx=0, pady=0, sticky="nwsw")

        # Configure center container/frame
        self.center_container_frame = customtkinter.CTkFrame(master=self, width=600, corner_radius=15, border_width=3,
                                                             border_color=('cyan','gray38'))
        self.center_container_frame.grid(row=0, column=1, padx=0, pady=20, sticky="ns")

        # Configure overlay between header and notes section
        self.center_container_frame_overlay1 = customtkinter.CTkFrame(master=self.center_container_frame, width=580,
                                                                      corner_radius=15,
                                                                      border_color=('white', 'gray38'),
                                                                      border_width=1)
        self.center_container_frame_overlay1.grid(row=2, column=1, padx=5, pady=10, sticky="nwsw")

        # Configure overlay between compatibility results and instruction label
        self.center_container_frame_overlay2 = customtkinter.CTkFrame(master=self.center_container_frame_overlay1,
                                                                      width=500, corner_radius=15,
                                                                      fg_color=('white', 'gray18'),
                                                                      border_color=('cyan'),
                                                                      border_width=1)
        self.center_container_frame_overlay2.grid(row=2, column=1, padx=5, pady=10, sticky="nw")

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
        self.compatibility_results_label = customtkinter.CTkLabel(master=self.center_container_frame_overlay1,
                                                                  text="Select Eagle Version From \nThe Drop Down Menu",
                                                                  height=80,
                                                                  fg_color=("white", "gray38"),
                                                                  text_font=('Verdana', 12))
        # Set the Eagle Version Label grid configuration
        self.compatibility_results_label.grid(row=1, column=0, columnspan=2, pady=20, padx=35, sticky='n')

        # Compatibility results label
        self.compatibility_results_label = customtkinter.CTkLabel(master=self.center_container_frame_overlay1,
                                                                  text="Compatibility results will \npopulate here.",
                                                                  height=50,
                                                                  fg_color=("white", "gray38"),
                                                                  text_font=('Verdana', 12))
        # Set the Compatibility results label grid configuration
        self.compatibility_results_label.grid(row=6, column=0, columnspan=2, pady=20, padx=5, sticky='n')

        # ============ Text Box Start ============

        # Text box
        self.display_exceptions = Text(master=self.center_container_frame, height=10, width=40, relief='sunken',
                                       bg='gray81', wrap='word', font=('Verdana', 9))
        self.display_exceptions.grid(row=7, column=0, columnspan=2, pady=0, padx=0, sticky='n')
        self.display_exceptions.insert('1.0',
                                       'Any additional compatibility notes i.e. special contingencies will display here.')

        # ============ Text Box End ============

        # Dummy label for spacing
        self.dummy_label = customtkinter.CTkLabel(master=self.center_container_frame,
                                                  text="")
        # Set the Compatibility results label grid configuration
        self.dummy_label.grid(row=8, column=0, columnspan=2, pady=5, padx=5, sticky='n')

        # ============ Labels End ============

        # ============ Option Menu Start ============

        # Eagle option menu
        self.eagle_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame_overlay2,
                                                                     values=self.eagle_versions_list,
                                                                     command=self.eagle_version_option_menu_callback)
        self.eagle_version_option_menu.grid(row=2, column=0, columnspan=2, pady=10, padx=45)

        # Software/device compatibility option menu
        self.software_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame_overlay2,
                                                                        values=self.software_device_list,
                                                                        command=self.device_software_menu_callback)
        self.software_version_option_menu.grid(row=4, column=0, columnspan=2, pady=10, padx=5)

        # ============ Combo Box(s) End ============

        # ============ Buttons Start ============

        self.radio_var = tk.IntVar()
        self.radio_var.set(0)

        # Toggle between oldest and newest eagle software versions
        self.reverse_eagle_version_radio_button = customtkinter.CTkRadioButton(
            master=self.center_container_frame_overlay2,
            text='toggle',
            command=self.reverse_eagle_version_list,
            variable=self.radio_var,
            value=0)
        self.reverse_eagle_version_radio_button.grid(row=3, column=0, columnspan=2, pady=10, padx=10, sticky="n")

        # Submit button
        self.submit_button = customtkinter.CTkButton(master=self.center_container_frame_overlay1,
                                                     text="Submit",
                                                     command=self.submit_choice_button,
                                                     border_width=2,
                                                     border_color='cyan')
        self.submit_button.grid(row=5, column=0, columnspan=2, pady=10, padx=5, sticky='n')

        # ============ Buttons End ============

        # ============ App Start Method Start ============

        # Read the Excel file into matrix upon app creation
        self.matrix_dictionary = self.get_matrix_dictionary()

        # ============ App Start Method End ==============

    # ============ Methods Start ============

    '''
        Method: on_closing
        Purpose: Destroy the entire application window when exit code = 0.
    '''

    def on_closing(self, event=0):
        return self.destroy()  # Destroy the entire application i.e. close out

    '''
         Method: option_menu_callback
         Purpose: Once the user selects a choice (Eagle version) from the drop down menu,
                  the global eagle version variable will be updated to reflect the user's current selection.
     '''

    def eagle_version_option_menu_callback(self, choice='V1002 – V1303'):
        self.global_eagle_version_choice = choice
        return choice  # return users Eagle version choice

    '''
         Method: device_software_menu_callback
         Purpose: Once the user selects a choice (software/device) from the drop down menu,
                  the global device/software variable will be updated to reflect the user's current selection.
     '''

    def device_software_menu_callback(self, choice='SafetyNet'):
        self.global_software_device_choice = choice
        return ''

    '''
        Method: get_device_index_from_list
        Purpose: To return the index of the user's device/software choice,
                 so it can be mapped to the compatibility list.
                 
                 - device: Contains the user's device/software choice i.e. 'SafetyNet', 'MICT', 'Sketch' etc..
    '''

    def get_device_index_from_list(self, device):
        return self.software_device_list.index(device)

    '''
        Method: open_excel_file
        Purpose: Open the Excel file, find the second table (header=16) located in the Excel sheet, 
                 and load columns B through Q (usecols='B:Q'), load rows (nrows=54) into a pandas dataframe.
    '''

    def open_excel_file(self):
        return pd.read_excel(self.file_path, header=16, usecols='B:Q', nrows=54)

    '''
        Method: get_columns
        Purpose: Simple helper method that returns the columns in a Pandas dataframe.
        
                - df: contains the Pandas dataframe
    '''

    def get_columns(self, df):
        return df.columns

    '''
        Method: convert_dataframe_to_numpy_array
        Purpose: Convert pandas dataframe to numpy array.  The numpy array is a list of lists
                 in the following format:
                 
                 [['V1002 – V1303' 'Yes[1]' 'Yes' 'No' 'No' 'No' 'No' 'No' 'No' 'No' 'No'
                   'Yes' 'No' 'No' 'No' 'No']]
        
                - df: contains pandas dataframe - which contains the entire excel table.
    '''

    def convert_dataframe_to_numpy_array(self, df):
        return df.to_numpy()

    '''
          Method: convert_numpy_array_to_dictionary
          Purpose: Convert numpy array which is a list of lists to a dictionary.
                   A dictionary will allow quick lookups by key.
                   
                   - array_np: contains the numpy array.
    '''

    def convert_numpy_array_to_dictionary(self, array_np):

        # Extract the eagle version as the key and compatibility matrix as the values
        for ls in array_np:
            # Grab eagle version as key
            key = ls[0]
            # Grab the rest of the list and use it as values
            values = ls[1:]
            # Set the key value pair
            self.matrix_dictionary[key] = values

        return self.matrix_dictionary  # Return the numpy matrix to caller

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

        return self.convert_numpy_array_to_dictionary(array_np)  # Return numpy array

    '''
         Method: find_exceptions_regex
         Purpose: to find pattern which contains one or two digits 
                  enclosed in brackets example [2] or [10]
    '''

    def find_exceptions_regex(self, string):
        # Find pattern which contains one or two digits enclosed in brackets example [2] or [10]
        pattern = self.get_pattern()

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

        # Everytime the reverse button is pressed, the first element in the eagle versions list will be the current
        # selected value i.e. version 'V1002 – V1303' or 'V2146'.
        self.global_eagle_version_choice = self.eagle_versions_list[0]

        # Display the updated Eagle option menu
        self.eagle_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame_overlay2,
                                                                     values=self.eagle_versions_list,
                                                                     command=self.eagle_version_option_menu_callback)
        self.eagle_version_option_menu.grid(row=2, column=0, columnspan=2, pady=10, padx=5)

        # When user presses 'reverse' button make sure software/device menu doesn't disappear.
        return self.software_device_menu()

    '''
          Method: software_device_menu
          Purpose: Display drop down menu that contains software/devices for root compatibility.
                   i.e. Safetynet, Iris, Sedline etc...
     '''

    def software_device_menu(self):
        # Software/device compatibility option menu
        self.software_version_option_menu = customtkinter.CTkOptionMenu(master=self.center_container_frame_overlay2,
                                                                        values=
                                                                        self.software_device_list[
                                                                            self.global_software_device_choice],
                                                                        command=self.device_software_menu_callback)
        return self.software_version_option_menu.grid(row=4, column=0, columnspan=2, pady=10, padx=5)

    '''
           Method: get_compatibility_list
           Purpose: Pass the eagle version the user chose - for example ('V1002 – V1303') - as 'key' to the dictionary 
                    that contains a list of corresponding 'yes' or 'no' values for example ['Yes','No'].
    '''

    def get_compatibility_list(self):
        return self.matrix_dictionary[self.global_eagle_version_choice]

    '''
        Method: is_compatible
        Purpose: To map the device/software to it's corresponding 'Yes' or 'No' value.
                 This method takes two parameters compatibility_list and device_choice.
                  
                - compatibility_list: contains a list of 'Yes' or 'No' values example ['Yes', 'No', 'Yes']
                  which maps to the software/device by index. 
                  
                - device_choice: contains 'SafetyNet', 'MICT', 'Sketch' etc.. selected by user.
    '''

    def is_compatible(self, compatibility_list, device_choice):
        return compatibility_list[self.get_device_index_from_list(device_choice)]  # Yes or No

    '''
           Method: get_matrix_dictionary
           Purpose: Returns a dictionary with eagle software as key and
                    compatibility list as values for all devices.
    '''

    def get_matrix_dictionary(self):
        # Get a matrix/dictionary of the Excel sheet
        return self.process_excel_file()

    '''
            Method: display_compatibility_results
            Purpose: Display compatibility results to the user.
    '''

    def display_compatibility_results(self, compatibility_result):
        # Destroy previous results before displaying new results
        self.compatibility_results_label.destroy()

        # Make sure to display the results with proper grammar
        # if for example the result is Yes[4] etc.. only take 'Yes' and exclude [4]
        if compatibility_result[0:3] == 'Yes':
            compatibility_result = "Yes"
            is_or_is_not_var = 'is'
        else:
            is_or_is_not_var = 'is not'

        # Compatibility results label
        self.compatibility_results_label = customtkinter.CTkLabel(master=self.center_container_frame_overlay1,
                                                                  text=compatibility_result + ': eagle version ' + self.global_eagle_version_choice + f'\n {is_or_is_not_var} compatible with ' + self.global_software_device_choice,
                                                                  height=50,
                                                                  fg_color=("white", "gray38"),
                                                                  text_font=('Verdana', 12))
        # Set the Compatibility results label grid configuration
        return self.compatibility_results_label.grid(row=6, column=0, columnspan=2, pady=20, padx=5, sticky='n')

    '''
            Method: display_exceptions_text_box
            Purpose: Display any special requirements to user in the notes text box.
    '''

    def display_exceptions_text_box(self, exception):
        self.delete_text_from_exceptions_text_box()  # Delete current text in text box
        self.insert_text_into_text_box(exception)  # Insert any special requirements into text box
        return self.update_text_box_with_new_text()  # Update text box the new text

    '''
            Method: delete_text_from_exceptions_text_box
            Purpose: Deletes text from the special notes text box.
     '''

    def delete_text_from_exceptions_text_box(self):
        return self.display_exceptions.delete('1.0', 'end')

    '''
             Method: insert_text_into_text_box
             Purpose: Insert new text into special notes text box.
    '''

    def insert_text_into_text_box(self, text):
        return self.display_exceptions.insert('end', text)

    '''
             Method: update_text_box_with_new_text
             Purpose: Update special notes text box with new text.
    '''

    def update_text_box_with_new_text(self):
        return self.display_exceptions.update()

    '''
             Method: get_pattern
             Purpose: Returns compiled regular expression that matches 
    '''

    def get_pattern(self):
        # Find pattern which contains one or two digits enclosed in brackets example [2] or [10]
        return re.compile(r'\[(\d){1,2}\]')

    '''
              Method: submit_choice_button
              Purpose: When the user has selected the desired eagle version and device,
                       the submit button will pass the eagle version and device to 
                       compatibility methods and display the results to the user.
     '''

    def submit_choice_button(self):

        # Pass the eagle version the user chose - for example ('V1002 – V1303') - as 'key' to the matrix/dictionary.
        # A list of 'yes' or 'no' values will be returned for example ['Yes','No'].
        compatibility_list = self.get_compatibility_list()

        # Yes or No value will be returned
        compatibility_result = self.is_compatible(compatibility_list, self.global_software_device_choice)

        # Display compatibility results to user
        self.display_compatibility_results(compatibility_result)

        # Find exceptions if they exist for the Eagle version
        eagle_version_exception = self.find_exceptions_regex(self.global_eagle_version_choice)

        # Find exceptions if they exist for the selected device
        device_version_exception = self.find_exceptions_regex(compatibility_result)

        # If the selected eagle version and device both do not have any special requirements
        if eagle_version_exception == '' and device_version_exception == '':
            self.display_exceptions_text_box(
                "There are no special requirements for " + self.global_eagle_version_choice + " and " + self.global_software_device_choice + " compatibility.")
        # Else if the selected device does have a special requirement and the eagle version does not
        elif device_version_exception != '' and eagle_version_exception == '':
            self.display_exceptions_text_box(device_version_exception)
        elif eagle_version_exception == '':
            self.display_exceptions_text_box(eagle_version_exception)
        # Else both the eagle version and device have special requirements
        else:
            self.display_exceptions_text_box(eagle_version_exception + "\n\n")
            self.insert_text_into_text_box(device_version_exception)

        return ''

    # ============ Methods End ============
