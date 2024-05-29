# Outlook VBA Script for Calendar Event Title Update Automation

This VBA script for Microsoft Outlook automates the process of updating the subject line of calendar items in a shared calendar. It ensures that the time of the event is included in the subject line and updates the subject line if the time of the event changes.


## Usage

### prequisites
1. **Ensure All Macros are Enabled**: Use the ["Enabling Macro Security"](#enabling-macro-security) section.

### setup
2. **Open Outlook VBA Editor**: Press `Alt + F11`.
2. **Insert a New Module**: Right-click on `Project1 (VbaProject.OTM)` > `Insert` > `Module`.
3. **Rename CalendarMailAddress**: Change it to your shared calendar's email address.
4. **Copy and Paste the Code**: Copy the script from "macro.txt" into the new module.
5. **Initialize the Handler**: Run `Initialize_handler` to set up the event handlers.

### additional steps for easy re-use
7. **Add to Quick Access Toolbar**: Use the ["Adding to the Quick Access Toolbar"](#adding-to-the-quick-access-toolbar) section.
6. **Re Run**: On each first use after logout, click the Run symbol on the quick access toolbar.

This script keeps the subject line of each calendar item in the shared calendar updated with the correct start and end times, ensuring consistency and accuracy.


## Additional Sections

### Enabling Macro Security

Before using VBA macros, you need to enable macros in Outlook:

1. **Open Outlook Options**: Go to `File` > `Options`.
2. **Trust Center**: Select `Trust Center` from the menu on the left, then click `Trust Center Settings`.
3. **Macro Settings**: In the Trust Center window, select `Macro Settings` and choose `Notifications for all macros` or `Enable all macros`. Click `OK` to save your changes.

### Adding to the Quick Access Toolbar

1. **Open Outlook Options**: Go to `File` > `Options`.
2. **Quick Access Toolbar**: In the Outlook Options window, select `Quick Access Toolbar`.
3. **Choose Commands**: From the `Choose commands from` dropdown, select `Macros`.
4. **Add Macro**: Find your macro (e.g., `Project1.Initialize_handler`), select it, and click `Add`.
5. **Rename Macro**: Optionally, rename the macro for better readability.
6. **Save Changes**: Click `OK` to save your changes and close the Outlook Options window.


## Technical Documentation

- **Constants and Global Variables**
  - `CalendarMailAddress`: Email address of the shared calendar. *required*
  - `MainSeparator`: Chosen character that separates time and main subject text.
  - `TimeSeperator`: Chosen character that separates start and end times.

- **Event and Handler Initialization**
  - `Initialize_handler`: Sets up the event handler to monitor items in the shared calendar.

- **Functions for Handling Time and Subject**
  - `get_item_time`: Formats start and end times of an item.
  - `get_subject_time`: Extracts time from the subject line.
  - `get_subject_text`: Extracts main text from the subject line.
  - `format_full_subject`: Combines formatted time and main subject text.
  - `is_time_changed`: Checks if the time in the subject line differs from the actual event time.

- **Handling Subject Changes**
  - `change_subject`: Updates the subject line to include the event time if it’s missing or if the time has changed.

- **Event Handlers for Item Addition and Change**
  - `myOlitems_ItemAdd`: Triggered when a new item is added to the calendar.
  - `myOlitems_ItemChange`: Triggered when an existing item is changed.
