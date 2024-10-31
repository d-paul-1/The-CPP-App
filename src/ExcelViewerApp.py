import customtkinter as ctk
import webview
import os

class ExcelViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel Viewer")
        self.geometry("800x600")

        # Button to open the viewer
        self.open_viewer_button = ctk.CTkButton(self, text="Open Excel Viewer", command=self.open_excel_viewer)
        self.open_viewer_button.pack(pady=20)

    def open_excel_viewer(self):
        # Create an HTML file with the Excel viewer code
        html_code = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel-like Editor</title>
    <link rel="stylesheet" href="https://handsontable.github.io/handsontable/dist/handsontable.full.min.css">
    <script src="https://handsontable.github.io/handsontable/dist/handsontable.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
    <style>
        #example {
            width: 600px;
            height: 300px;
            overflow: hidden;
            margin: 20px;
            border: 1px solid #ccc;
        }
    </style>
</head>
<body>
    <h1>Excel-like Editor</h1>
    <input type="file" id="file" />
    <button id="download">Download Modified Excel</button>
    <div id="example"></div>
    <script>
        let hot;

        document.getElementById('file').addEventListener('change', (event) => {
            const file = event.target.files[0];
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    // Convert worksheet to JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // Create Handsontable instance
                    hot = new Handsontable(document.getElementById('example'), {
                        data: jsonData,
                        colHeaders: true,
                        rowHeaders: true,
                        filters: true,
                        dropdownMenu: true,
                        contextMenu: true,
                        licenseKey: 'non-commercial-and-evaluation' // Handsontable license key
                    });
                } catch (error) {
                    alert('Error reading the Excel file. Please make sure it is a valid .xlsx file.');
                    console.error(error);
                }
            };

            reader.onerror = (error) => {
                alert('Error reading file. Please try again.');
                console.error(error);
            };

            reader.readAsArrayBuffer(file);
        });

        document.getElementById('download').addEventListener('click', () => {
            if (hot) {
                const modifiedData = hot.getData(); // Get data from Handsontable
                const newWorkbook = XLSX.utils.book_new();
                const newWorksheet = XLSX.utils.aoa_to_sheet(modifiedData);
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");
                XLSX.writeFile(newWorkbook, 'modified_excel.xlsx');
            } else {
                alert('No data to download. Please upload an Excel file first.');
            }
        });
    </script>
</body>
</html>
'''
        # Save the HTML file
        html_file_path = os.path.join(os.path.dirname(__file__), 'excel_viewer.html')
        with open(html_file_path, 'w') as html_file:
            html_file.write(html_code)

        # Create a web view window
        webview.create_window("Excel Viewer", html_file_path)
        webview.start()

if __name__ == "__main__":
    app = ExcelViewerApp()
    app.mainloop()
