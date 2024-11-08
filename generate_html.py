import pandas as pd

# Define the file path
file_path = 'IncidentPractice.xlsx'
output_file_path = 'output.html'

# Read the Excel file
xls = pd.ExcelFile(file_path, engine='xlrd')

# Custom sheet names
sheet_names = ["MIM Handover at 7 30 GMT", "MIM Handover at 16 30 GMT", "Process updates"]

# Generate HTML content
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Sheets as Tabs</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f9;
            margin: 0;
            padding: 20px;
        }
        .tab-buttons {
            display: flex;
            cursor: pointer;
            margin-bottom: 20px;
        }
        .tab-buttons div {
            padding: 10px 20px;
            border: 1px solid #ccc;
            margin-right: 5px;
            background-color: #e0e0e0;
            border-radius: 5px;
        }
        .tab-buttons .active {
            background-color: #007bff;
            color: white;
        }
        .tab {
            display: none;
        }
        .tab.active {
            display: block;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        table, th, td {
            border: 1px solid #ddd;
        }
        th, td {
            padding: 10px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        h2 {
            color: #333;
        }
        .spacer {
            height: 20px;
        }
    </style>
</head>
<body>
    <div class="tab-buttons" id="tab-buttons">"""

for index, sheet_name in enumerate(xls.sheet_names[:2]):
    custom_name = sheet_names[index] if index < len(sheet_names) else sheet_name
    html_content += '<div class="tab-button" onclick="showTab({})">{}</div>'.format(index, custom_name)

html_content += '</div><div id="tabs">'

# Read the process updates sheet
process_updates_df = pd.read_excel(xls, xls.sheet_names[2], usecols="B:F")

for index, sheet_name in enumerate(xls.sheet_names[:2]):
    df = pd.read_excel(xls, sheet_name, usecols="B:G")  # Read from column two only (B to G)
    html_content += '<div id="tab-{}" class="tab {}">'.format(index, 'active' if index == 0 else '')
    
    major_incidents = df[df.iloc[:, 5] == 'Major']
    p2_incidents = df[df.iloc[:, 5] == 'P2']
    
    if not major_incidents.empty:
        html_content += '<h2>Major</h2>'
        for _, row in major_incidents.iterrows():
            html_content += """
            <table>
                <tr><th>Incident:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Description:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Impact:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Current Status:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>ETA for Resolution:</th><td><input type="text" value="{}" readonly></td></tr>
            </table>
            <div class="spacer"></div>
            """.format(row[0], row[1], row[2], row[3], row[4])
    
    if not p2_incidents.empty:
        html_content += '<h2>P2</h2>'
        for _, row in p2_incidents.iterrows():
            html_content += """
            <table>
                <tr><th>Incident:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Description:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Impact:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>Current Status:</th><td><input type="text" value="{}" readonly></td></tr>
                <tr><th>ETA for Resolution:</th><td><input type="text" value="{}" readonly></td></tr>
            </table>
            <div class="spacer"></div>
            """.format(row[0], row[1], row[2], row[3], row[4])

    # Add process updates
    if index == 0 and len(process_updates_df) > 1:
        process_update_row = process_updates_df.iloc[1]
        html_content += '<h2>Process Updates</h2>'
        html_content += """
        <table>
            <tr><th>Introduced on:</th><td><input type="text" value="{}" readonly></td></tr>
            <tr><th>Current Process Updates:</th><td><input type="text" value="{}" readonly></td></tr>
        </table>
        <div class="spacer"></div>
        """.format(process_update_row[0], process_update_row[1])
    
    if index == 1 and len(process_updates_df) > 2:
        process_update_row = process_updates_df.iloc[2]
        html_content += '<h2>Process Updates</h2>'
        html_content += """
        <table>
            <tr><th>Introduced on:</th><td><input type="text" value="{}" readonly></td></tr>
            <tr><th>Current Process Updates:</th><td><input type="text" value="{}" readonly></td></tr>
        </table>
        <div class="spacer"></div>
        """.format(process_update_row[0], process_update_row[1])

    html_content += '</div>'

html_content += """
    </div>
    <script>
        function showTab(index) {
            var tabs = document.querySelectorAll('.tab');
            var buttons = document.querySelectorAll('.tab-button');
            tabs.forEach(function(tab, i) {
                tab.classList.toggle('active', i === index);
                buttons[i].classList.toggle('active', i === index);
            });
        }
    </script>
</body>
</html>"""

# Write the HTML content to a file
with open(output_file_path, 'w') as file:
    file.write(html_content)

print('HTML file generated at {}'.format(output_file_path))
