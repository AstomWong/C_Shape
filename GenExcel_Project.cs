using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;



namespace GenExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            Excel.Application OrgWorkExcel;
            Excel.Application NewWorkExcel;
            Excel.Workbook OrgWorkBook;
            Excel.Workbook NewWorkBook;
            Excel.Worksheet NewWorkSheet;
            Excel.Worksheet OrgWorkSheet;
            
            Excel.Range xlCell;
            Excel.Range oRange;
            Excel.Range Data;
            Excel.Range otherData;
            Excel.Range WholeCell;
            Excel.Range BorderRow;
            Excel.Borders BorderOfCell;


            int xCol = 1;
            int xRow = 1;
            int RowCount;
            int ColCount;
            int Num = 1;

            String SignID;
            String New_Building;
            String New_Level;
            String New_Zone;
            String Old_Building = "none";
            String Old_Level = "none";
            String Old_Zone = "none";
            String mData;
            String IMG_ID;
            String[] Title = new String[9];



            OrgWorkExcel = new Excel.Application();
            OrgWorkBook = OrgWorkExcel.Workbooks.Open(@"E:\Airport_Project\2010511\GenExcel\MFC_Sign_Id.xlsx");      // Open the Original Workbook
            OrgWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)OrgWorkBook.Sheets["Sheet1"];                  // Open the Original Workbooks' sheet 1

            NewWorkExcel = new Excel.Application();
            NewWorkBook = NewWorkExcel.Workbooks.Add();                                                             //Create a new excel file
            NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkBook.Sheets["Sheet1"];                  //Open the new excel sheet 1
            WholeCell = NewWorkSheet.get_Range("A1", "AZ10000");
            otherData = NewWorkSheet.get_Range("A2", "AZ10000");
            otherData.RowHeight = 220;

            //Set global attributes
            NewWorkExcel.StandardFont = "Gulim";                                                                    //This line is set up the Font
            WholeCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            RowCount = OrgWorkSheet.UsedRange.Rows.Count;                                                           //Get the total row of the excel
            ColCount = OrgWorkSheet.UsedRange.Columns.Count;
            //System.Diagnostics.Debug.Print(Convert.ToString(RowCount));
            // PreSet the 
            Title[0] = "Sign ID";
            Title[1] = "Side A Type_ID";
            Title[2] = "Side B Type_ID";
            Title[3] = "Changed";
            Title[4] = "IMG_Side A";
            Title[5] = "IMG_Side B";
            Title[6] = "Building";
            Title[7] = "Level";
            Title[8] = "Zone";

            for (int iRow = 1; iRow <= RowCount; iRow++)
            {

                if (xRow == 1)
                {

                    System.Diagnostics.Debug.Print(Convert.ToString(xRow));

                    for (int iCol = 0; iCol <= 8; iCol++)
                    {
                        NewWorkSheet.Cells[xRow, iCol + 1] = Title[iCol];
                    }

                    xRow++;

                }
                else if (xRow > 1)
                {
                        
                    xlCell = (Excel.Range)OrgWorkSheet.Cells[iRow, 1];

                    SignID = xlCell.Value.ToString();

                    string[] words = SignID.Split('/');

                    New_Building = words[0];
                    New_Level = words[1];
                    New_Zone = words[2];

                    //Compare building's part
                    if (New_Building != Old_Building)
                    {
                        NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\"+Old_Building+"_"+Old_Level + "_" + Old_Zone+".xls", Excel.XlFileFormat.xlWorkbookNormal);
                        NewWorkBook.Close(true);

                        NewWorkExcel = new Excel.Application();
                        NewWorkBook = NewWorkExcel.Workbooks.Add();                                                             //Create a new excel file
                        NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkBook.Sheets["Sheet1"];                  //Open the new excel sheet 1
                        WholeCell = NewWorkSheet.get_Range("A1", "AZ10000");
                        otherData = NewWorkSheet.get_Range("A2", "AZ10000");
                        otherData.RowHeight = 220;

                        //Set global attributes
                        NewWorkExcel.StandardFont = "Gulim";
                        WholeCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        for (int iCol = 0; iCol <= 8; iCol++)
                        {
                            NewWorkSheet.Cells[1, iCol + 1] = Title[iCol];
                        }

                        Old_Building = New_Building;
                        Old_Level = New_Level;
                        Old_Zone = New_Zone;
                        xRow = 2;
                    }
                    else if (New_Level != Old_Level)
                    {
                        NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\" + Old_Building + "_" + Old_Level + "_" + Old_Zone + ".xls", Excel.XlFileFormat.xlWorkbookNormal);
                        NewWorkBook.Close(true);

                        NewWorkExcel = new Excel.Application();
                        NewWorkBook = NewWorkExcel.Workbooks.Add();                                                             //Create a new excel file
                        NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkBook.Sheets["Sheet1"];                  //Open the new excel sheet 1
                        WholeCell = NewWorkSheet.get_Range("A1", "AZ10000");
                        otherData = NewWorkSheet.get_Range("A2", "AZ10000");
                        otherData.RowHeight = 220;

                        //Set global attributes
                        NewWorkExcel.StandardFont = "Gulim";
                        WholeCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        for (int iCol = 0; iCol <= 8; iCol++)
                        {
                            NewWorkSheet.Cells[1, iCol + 1] = Title[iCol];
                        }

                        Old_Level = New_Level;
                        Old_Zone = New_Zone;
                        xRow = 2;
                    }
                    else if (New_Zone != Old_Zone)
                    {
                        NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\" + Old_Building + "_" + Old_Level + "_" + Old_Zone + ".xls", Excel.XlFileFormat.xlWorkbookNormal);
                        NewWorkBook.Close(true);

                        NewWorkExcel = new Excel.Application();
                        NewWorkBook = NewWorkExcel.Workbooks.Add();                                                             //Create a new excel file
                        NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkBook.Sheets["Sheet1"];                  //Open the new excel sheet 1
                        WholeCell = NewWorkSheet.get_Range("A1", "AZ10000");
                        otherData = NewWorkSheet.get_Range("A2", "AZ10000");
                        otherData.RowHeight = 220;

                        //Set global attributes
                        NewWorkExcel.StandardFont = "Gulim";
                        WholeCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                        for (int iCol = 0; iCol <= 8; iCol++)
                        {
                            NewWorkSheet.Cells[1, iCol + 1] = Title[iCol];
                        }

                        Old_Zone = New_Zone;
                        xRow = 2;
                    }

                    //System.Diagnostics.Debug.Print("Building: " + words[0] +"/"+words[1]+"/"+words[2]);
                    //System.Diagnostics.Debug.Print("New Building: " + Old_Building + "/"+ Old_Level + "/"+ Old_Zone);
                    int DataLoop = 0;
                    for (int iCol = 1; iCol <= ColCount; iCol++)
                    {
                        if (iCol == 1 || iCol == 2 || iCol == 4)
                        {
                            
                            NewWorkSheet.Cells[xRow, xCol] = OrgWorkSheet.Cells[iRow, iCol];
                            
                            if ( iCol == 2 || iCol == 4)
                            {
                                int uCol = xCol + 3;
                                System.Diagnostics.Debug.Print("uCol: " + uCol);
                                Data = (Excel.Range)OrgWorkSheet.Cells[iRow, iCol];
                                oRange = (Excel.Range)NewWorkSheet.Cells[xRow, uCol];

                                float Top = (float)((double)oRange.Top);
                                float Left = (float)((double)oRange.Left);
                                int Width = 42;
                                int Height = 220;
                                oRange.RowHeight = Height;
                                oRange.ColumnWidth = Width;

                                const float ImageSize = 30;

                                if (Data.Value != null) { 
                                    //System.Diagnostics.Debug.Print("DataLoop: " + Convert.ToString(DataLoop));
                                    mData = Data.Value.ToString();
                                    if (mData.Split('-') != null)
                                    {
                                        string[] mWords = mData.Split('-');
                                        IMG_ID = "DSC0" + mWords[1] + ".jpg";
                                        NewWorkSheet.Cells[xRow, xCol] = IMG_ID;
                                        System.Diagnostics.Debug.Print("Image ID: " + IMG_ID);
                                        NewWorkSheet.Shapes.AddPicture(@"E:\Airport_Project\20170516\QA_IMG_PlatForm\Resize_Image\" + IMG_ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, 255, Height);
                                        //NewWorkSheet.Shapes.AddPicture(@"E:\Airport_Project\20170509\QA_IMG_PlatForm\QA_IMG_Collected_Resize\" + mData[DataLoop] + ".JPG", MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, Width, Height);
                                        DataLoop++;
                                    }
                                    else if (mData.Split('_') != null)
                                    {
                                        string[] mWords = mData.Split('-');
                                        IMG_ID = "IMG_" + mWords[1] + ".jpg";
                                        NewWorkSheet.Cells[xRow, xCol] = IMG_ID;
                                        System.Diagnostics.Debug.Print("Image ID: " + IMG_ID);
                                        NewWorkSheet.Shapes.AddPicture(@"E:\Airport_Project\20170516\QA_IMG_PlatForm\Resize_Image\" + IMG_ID, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, 255, Height);
                                        //NewWorkSheet.Shapes.AddPicture(@"E:\Airport_Project\20170509\QA_IMG_PlatForm\QA_IMG_Collected_Resize\" + mData[DataLoop] + ".JPG", MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, Width, Height);
                                        DataLoop++;
                                    }



                                }

                            }
                            xCol++;
                        }
                        else if (iCol == 6 || iCol == 7 || iCol == 8)
                        {

                            xCol = iCol + 1;
                            NewWorkSheet.Cells[xRow, xCol] = OrgWorkSheet.Cells[iRow, iCol];
                            //Setup the formate of the Cells
                            Excel.Range Location = (Microsoft.Office.Interop.Excel.Range)NewWorkSheet.Cells[xRow, xCol];
                            Location.Font.Bold = true;
                            xCol = 1;
                            //NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\"+ Num +".xls", Excel.XlFileFormat.xlWorkbookNormal);
                            //System.Diagnostics.Debug.Print("Saving!");
                            Num++;
                        }
                    }
                    
                }
                xRow++;
            }

            //NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\Demo.xls", Excel.XlFileFormat.xlWorkbookNormal);
            NewWorkBook.SaveAs(@"E:\Airport_Project\2010511\GenExcel\" + Old_Building + "_" + Old_Level + "_" + Old_Zone + ".xls", Excel.XlFileFormat.xlWorkbookNormal);
            NewWorkBook.Close(true);
            OrgWorkBook.Close(true);
            MessageBox.Show("File created !");

        }
    }
}
