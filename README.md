# powerpoint-automation

**GitHub Description:**

📝 **PowerPoint Presentation Generator**

This Python script automates the generation of PowerPoint presentations for tracking shipment deliveries. It utilizes the popular `PySimpleGUI` library for creating a user-friendly interface to input the required data, and `pptx` to dynamically create and format slides with the relevant information.

**Features:**
- Simple GUI to input delivery details: date, time, delivering company, country, and computer address.
- Automatically generates a PowerPoint presentation with the provided information.
- Customizable slide layout and formatting to suit your preferences.
- Supports appending to an existing presentation or creating a new one.

**How to Use:**
1. Run the script.
2. Enter the delivery details in the GUI form.
3. Click on the "Run" button to generate the PowerPoint presentation.

**Note:** The script will automatically overwrite the existing presentation with the same name if one is found in the specified directory.

**Requirements:**
- Python 3.x
- PySimpleGUI
- python-pptx

**Disclaimer:**
The code includes commented out sections intended for local testing (process termination). Use them responsibly and remove them when deploying the code in production.
