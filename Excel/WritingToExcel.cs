    class WritingToExcel
    {
        static void Main(string[] args)
        {
        // Get the app
        Excel.Application app = new Excel.Application();
        
        // Run the process in hidden state
        app.Visible = false;
        
        // Open the workbook with given path
        Excel.Workbook workbook = app.Workbooks.Open(@"C:\Users\Maicon Rossini\Desktop\PASTA TESTE\arq.xlsx",0,false);
        
        //Open the worksheet with hardcoded integer number
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
            
            // Just insert and test
            worksheet.Cells[1, 1] = 25;
            worksheet.Cells[1, 2] = 35;
            worksheet.Cells[1, 3] = 45;
            
            // Save the workbook
            workbook.Save();
            
            // Close the workbook
            workbook.Close(1);
            
            // Close the application - if it not got closed the upcoming open of this file will be readonly, pay attention
            app.Quit();
        }
    }
