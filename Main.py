import PySimpleGUI as sg
from backend.create_dashboard import Create_Dashboard

# Define the layout GUI
layout = [
    [sg.Text("Dashboard Width (Int)"), sg.InputText(key="Dashboard Width (Int)")],
    [sg.Text("Dashboard Length (Int)"), sg.InputText(key="Dashboard Length (Int)")],
    [sg.Text("Metric Count"), sg.InputText(key="Metric Count")],
    [sg.Text("Metric Minimum Width (Int)"), sg.InputText(key="Metric Minimum Width (Int)")],
    [sg.Text("Metric Minimum Length (Int)"), sg.InputText(key="Metric Minimum Length (Int)")],
    [sg.Button("Run")],
    [sg.Output(size=(60, 10))]
]
# Create the window
window = sg.Window("Metric Calculator", layout)

# Event loop
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break
    elif event == "Run":
        metric1 = values["Dashboard Width (Int)"]
        metric2 = values["Dashboard Length (Int)"]
        metric3 = values["Metric Count"]
        metric4 = values["Metric Minimum Width (Int)"]
        metric5 = values["Metric Minimum Length (Int)"]

        try:
            #Create a verify Method
            int(metric1)
            int(metric2)
            int(metric3)
            int(metric4)
            int(metric5)
        except:
            print("Please Only Enter Intergers")
            print("Example Values:")
            print("Height = 50")
            print("Width = 80")
            print("Metric Count = 6")
            print("Min Metric Width = 6")
            print("Min Metric Height = 6")
            
        else:
            Create_Dashboard(metric1, metric2,metric3,metric4,metric5)

# Close the window
window.close()


