using IronXL;
using SixLabors.ImageSharp.PixelFormats;

namespace MAUI_IronXL;

public partial class MainPage : ContentPage
{
	public MainPage()
	{
		InitializeComponent();
	}

	private void CreateExcel(object sender, EventArgs e)
	{
        //Create Workbook
        WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
        
        //Create Worksheet
        var sheet = workbook.CreateWorkSheet("2022 Budget");
        
        //Set Cell values
        sheet["A1"].Value = "January";
        sheet["B1"].Value = "February";
        sheet["C1"].Value = "March";
        sheet["D1"].Value = "April";
        sheet["E1"].Value = "May";
        sheet["F1"].Value = "June";
        sheet["G1"].Value = "July";
        sheet["H1"].Value = "August";

        //Set Cell values Dynamically
        Random r = new();
        for (int i = 2; i <= 11; i++)
        {
            sheet["A" + i].Value = r.Next(1, 1000);
            sheet["B" + i].Value = r.Next(1000, 2000);
            sheet["C" + i].Value = r.Next(2000, 3000);
            sheet["D" + i].Value = r.Next(3000, 4000);
            sheet["E" + i].Value = r.Next(4000, 5000);
            sheet["F" + i].Value = r.Next(5000, 6000);
            sheet["G" + i].Value = r.Next(6000, 7000);
            sheet["H" + i].Value = r.Next(7000, 8000);
        }

        //Apply formatting like background and border
        sheet["A1:H1"].Style.SetBackgroundColor("#d3d3d3");
        sheet["A1:H1"].Style.TopBorder.SetColor("#000000");
        sheet["A1:H1"].Style.BottomBorder.SetColor("#000000");
        sheet["H2:H11"].Style.RightBorder.SetColor("#000000");
        sheet["H2:H11"].Style.RightBorder.Type = IronXL.Styles.BorderType.Medium;
        sheet["A11:H11"].Style.BottomBorder.SetColor("#000000");
        sheet["A11:H11"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Medium;

        //Apply Formulas
        decimal sum = sheet["A2:A11"].Sum();
        decimal avg = sheet["B2:B11"].Avg();
        decimal max = sheet["C2:C11"].Max();
        decimal min = sheet["D2:D11"].Min();

        sheet["A12"].Value = "Sum";
        sheet["B12"].Value = sum;


        sheet["C12"].Value = "Avg";
        sheet["D12"].Value = avg;

        sheet["E12"].Value = "Max";
        sheet["F12"].Value = max;

        sheet["G12"].Value = "Min";
        sheet["H12"].Value = min;


        //Save and Open Excel File
        SaveService saveService = new();
        saveService.SaveAndView("Budget.xlsx", "application/octet-stream", workbook.ToStream());
    }

    private void ReadExcel(object sender, EventArgs e)
    {
        WorkBook workbook = WorkBook.Load(@"C:\Files\Customer Data.xlsx");
        WorkSheet sheet = workbook.WorkSheets.First();

        decimal sum = sheet["B2:B10"].Sum();

        sheet["B11"].Value = sum;
        sheet["B11"].Style.SetBackgroundColor("#808080");
        sheet["B11"].Style.Font.SetColor("#ffffff");

        //Save and Open Excel File
        SaveService saveService = new();
        saveService.SaveAndView("Modified Data.xlsx", "application/octet-stream", workbook.ToStream());

        DisplayAlert("Notification", "Excel file has been modified!", "OK");
    }
}

