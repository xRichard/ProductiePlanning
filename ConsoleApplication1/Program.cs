using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using ClosedXML;

namespace ProductionPlanning
{
    class Program
    {
        static void Main(string[] args)
        {

            //Still ToDO:
            //Split everything up based on divisions

            //Add the e-mailing section

            //Check if people need the ability to see more then one day.


            //Setting the connection string
            string conString = "Provider = IBMDA400; Data Source = 192.168.250.4; User Id = AURORA;Password = AURORA;";
            OleDbConnection conn = new OleDbConnection(conString);
            //Opening the connection
            conn.Open();
            
            DateTime today = DateTime.Today;
            string dateToday = "1" + today.ToString("yyMMdd");
            string QueryProdPlanToday = "SELECT ORDN55, CUSN55, CATN55, QTOR55, EUOM55, TOWT55, PDES35, CNAM05, CAD305, DSEQ55, DIVN55 FROM AULT2F2.OEP55 LEFT OUTER JOIN AULT1F2.SLP05 ON CONO05 = CONO55 AND CUSN05 = CUSN55 AND DSEQ05 = DSEQ55 LEFT OUTER JOIN AULT2F2.INP35 ON CONO35 = CONO55 AND PNUM35 = CATN55 WHERE CONO55 = '21' AND DTDR55 = '" + dateToday + "' AND EUOM55 != 'ST' ORDER BY DIVN55, CATN55";
            OleDbCommand cmd = new OleDbCommand(QueryProdPlanToday);
            cmd.Connection = conn;
            cmd.CommandType = CommandType.Text;

            var wb = new ClosedXML.Excel.XLWorkbook();
            var wsToday = wb.Worksheets.Add("Prod_Vandaag");
            var wsTom = wb.Worksheets.Add("Prod Tomorrow");

            //Setting page orientation
            wsToday.PageSetup.PageOrientation   =   ClosedXML.Excel.XLPageOrientation.Landscape;
            wsTom.PageSetup.PageOrientation     =   ClosedXML.Excel.XLPageOrientation.Landscape;
            
            //Building the excel header for today
            wsToday.Cell("A1").Value        =       "Art Nr";
            wsToday.Cell("B1").Value        =       "Art Desc";
            wsToday.Cell("C1").Value        =       "Quantity ordered";
            wsToday.Cell("D1").Value        =       "Total weight";
            wsToday.Cell("E1").Value        =       "Packing";
            wsToday.Cell("F1").Value        =       "Customer name";
            wsToday.Cell("G1").Value        =       "Customer address";
            wsToday.Cell("H1").Value        =       "Customer nr";          
            wsToday.Cell("I1").Value        =       "Cust Delivery code";

            wsToday.Cell("A1").Style.Font.Bold = true;
            wsToday.Cell("B1").Style.Font.Bold = true;
            wsToday.Cell("C1").Style.Font.Bold = true;
            wsToday.Cell("D1").Style.Font.Bold = true;
            wsToday.Cell("E1").Style.Font.Bold = true;
            wsToday.Cell("F1").Style.Font.Bold = true;
            wsToday.Cell("G1").Style.Font.Bold = true;
            wsToday.Cell("H1").Style.Font.Bold = true;
            wsToday.Cell("I1").Style.Font.Bold = true;
            
            //Building the excel header for tomorrow
            wsTom.Cell("A1").Value = "Art Nr";
            wsTom.Cell("B1").Value = "Art Desc";
            wsTom.Cell("C1").Value = "Quantity ordered";
            wsTom.Cell("D1").Value = "Total weight";
            wsTom.Cell("E1").Value = "Packing";
            wsTom.Cell("F1").Value = "Customer name";
            wsTom.Cell("G1").Value = "Customer address";
            wsTom.Cell("H1").Value = "Customer nr";
            wsTom.Cell("I1").Value = "Cust Delivery code";

            wsTom.Cell("A1").Style.Font.Bold = true;
            wsTom.Cell("B1").Style.Font.Bold = true;
            wsTom.Cell("C1").Style.Font.Bold = true;
            wsTom.Cell("D1").Style.Font.Bold = true;
            wsTom.Cell("E1").Style.Font.Bold = true;
            wsTom.Cell("F1").Style.Font.Bold = true;
            wsTom.Cell("G1").Style.Font.Bold = true;
            wsTom.Cell("H1").Style.Font.Bold = true;
            wsTom.Cell("I1").Style.Font.Bold = true;

            //Setting this int so we have a baseline on where to start writing the lines in the excel file
            int a = 2;

            try
            {
                OleDbDataReader reader = cmd.ExecuteReader();

                //section of today
                while (reader.Read())
                {
                    //Defining the vars and making sure they are empty
                    string OrdNr = "";
                    string CustNr = "";
                    string ArtNr = "";
                    string QuantityOrdered = "";
                    string Packing = "";
                    string TotalWeight = "";
                    string ArtDescr = "";
                    string CustName = "";
                    string CustAddr = "";
                    string CustDelCode = "";

                    //The fields we select from the database:
                    //  0   =   ORDN55
                    //  1   =   CUSN55
                    //  2   =   CATN55
                    //  3   =   QTOR55
                    //  4   =   EUOM55
                    //  5   =   TOWT55
                    //  6   =   PDES35
                    //  7   =   CNAM05
                    //  8   =   CAD305
                    //  9   =   DSEQ55

                    //Ordernumber
                    if (!reader.IsDBNull(0))
                    {
                        OrdNr = reader.GetValue(0).ToString();
                    }
                    //CustomerNumber
                    if (!reader.IsDBNull(1))
                    {
                        CustNr = reader.GetValue(1).ToString();
                    }
                    //Product number
                    if (!reader.IsDBNull(2))
                    {
                        ArtNr = reader.GetValue(2).ToString();
                    }
                    //Quantity ordered - apparently this is the only thing they need to decide on how much they wanna produce....
                    if (!reader.IsDBNull(3))
                    {
                        QuantityOrdered = reader.GetValue(3).ToString();
                    }
                    //Package material, like box, crate, etc
                    if (!reader.IsDBNull(4))
                    {
                        Packing = reader.GetValue(4).ToString();
                    }
                    //Total physical weight
                    if (!reader.IsDBNull(5))
                    {
                        TotalWeight = reader.GetValue(5).ToString();
                    }
                    //Article description - name of the product
                    if (!reader.IsDBNull(6))
                    {
                        ArtDescr = reader.GetValue(6).ToString();
                    }
                    //Customer name
                    if (!reader.IsDBNull(7))
                    {
                        CustName = reader.GetValue(7).ToString();
                    }
                    //Customer address/country
                    if (!reader.IsDBNull(8))
                    {
                        CustAddr = reader.GetValue(8).ToString();
                    }
                    //Customer delivery code
                    if(!reader.IsDBNull(9))
                    {
                        CustDelCode = reader.GetValue(9).ToString();
                    }
                    
                    //Debugging only
                    Console.WriteLine(OrdNr);

                    //Filled the vars - now just gotta build up the excel file

                    //Filling the excel cells
                    wsToday.Cell(a, 1).Value = ArtNr;
                    wsToday.Cell(a, 2).Value = ArtDescr;
                    wsToday.Cell(a, 3).Value = QuantityOrdered;
                    wsToday.Cell(a, 4).Value = TotalWeight;
                    wsToday.Cell(a, 5).Value = Packing;
                    wsToday.Cell(a, 6).Value = CustName;
                    wsToday.Cell(a, 7).Value = CustAddr;
                    wsToday.Cell(a, 8).Value = CustNr;
                    wsToday.Cell(a, 9).Value = CustDelCode;

                    //Setting the horizontal alignment to left on the fields with decimal number in there so it looks a bit better.
                    wsToday.Cell(a, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsToday.Cell(a, 3).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsToday.Cell(a, 4).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsToday.Cell(a, 6).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsToday.Cell(a, 8).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsToday.Cell(a, 9).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                                        
                    //Raising a so that all the lines will actually be inserted
                    a++;


                }//End of while loop
                Console.WriteLine("Starting the 2nd worksheet now");

                DateTime tomorrow = DateTime.Today.AddDays(1);
                string dateTom = "1" + tomorrow.ToString("yyMMdd");
            

                string QueryProdPlanTom = "SELECT ORDN55, CUSN55, CATN55, QTOR55, EUOM55, TOWT55, PDES35, CNAM05, CAD305, DSEQ55, DIVN55 FROM AULT2F2.OEP55 LEFT OUTER JOIN AULT1F2.SLP05 ON CONO05 = CONO55 AND CUSN05 = CUSN55 AND DSEQ05 = DSEQ55 LEFT OUTER JOIN AULT2F2.INP35 ON CONO35 = CONO55 AND PNUM35 = CATN55 WHERE CONO55 = '21' AND DTDR55 = '" + dateTom + "' AND EUOM55 != 'ST' ORDER BY DIVN55, CATN55";
                OleDbCommand cmdTom = new OleDbCommand(QueryProdPlanTom);
                cmdTom.Connection = conn;
                cmdTom.CommandType = CommandType.Text;

                //Setting the int for tomorrow sheet
                int b = 2;

                //Section of tomorrow
                OleDbDataReader readerTom = cmdTom.ExecuteReader();
                while (readerTom.Read())
                {
                    //Defining the vars and making sure they are empty
                    string OrdNr = "";
                    string CustNr = "";
                    string ArtNr = "";
                    string QuantityOrdered = "";
                    string Packing = "";
                    string TotalWeight = "";
                    string ArtDescr = "";
                    string CustName = "";
                    string CustAddr = "";
                    string CustDelCode = "";

                    //Ordernumber
                    if (!readerTom.IsDBNull(0))
                    {
                        OrdNr = readerTom.GetValue(0).ToString();
                    }
                    //CustomerNumber
                    if (!readerTom.IsDBNull(1))
                    {
                        CustNr = readerTom.GetValue(1).ToString();
                    }
                    //Product number
                    if (!readerTom.IsDBNull(2))
                    {
                        ArtNr = readerTom.GetValue(2).ToString();
                    }
                    //Quantity ordered - apparently this is the only thing they need to decide on how much they wanna produce....
                    if (!readerTom.IsDBNull(3))
                    {
                        QuantityOrdered = readerTom.GetValue(3).ToString();
                    }
                    //Package material, like box, crate, etc
                    if (!readerTom.IsDBNull(4))
                    {
                        Packing = readerTom.GetValue(4).ToString();
                    }
                    //Total physical weight
                    if (!readerTom.IsDBNull(5))
                    {
                        TotalWeight = readerTom.GetValue(5).ToString();
                    }
                    //Article description - name of the product
                    if (!readerTom.IsDBNull(6))
                    {
                        ArtDescr = readerTom.GetValue(6).ToString();
                    }
                    //Customer name
                    if (!readerTom.IsDBNull(7))
                    {
                        CustName = readerTom.GetValue(7).ToString();
                    }
                    //Customer address/country
                    if (!readerTom.IsDBNull(8))
                    {
                        CustAddr = readerTom.GetValue(8).ToString();
                    }
                    //Customer delivery code
                    if (!readerTom.IsDBNull(9))
                    {
                        CustDelCode = readerTom.GetValue(9).ToString();
                    }

                    //Debugging only
                    Console.WriteLine(OrdNr);

                    //Filled the vars - now just gotta build up the excel file

                    //Filling the excel cells
                    wsTom.Cell(b, 1).Value = ArtNr;
                    wsTom.Cell(b, 2).Value = ArtDescr;
                    wsTom.Cell(b, 3).Value = QuantityOrdered;
                    wsTom.Cell(b, 4).Value = TotalWeight;
                    wsTom.Cell(b, 5).Value = Packing;
                    wsTom.Cell(b, 6).Value = CustName;
                    wsTom.Cell(b, 7).Value = CustAddr;
                    wsTom.Cell(b, 8).Value = CustNr;
                    wsTom.Cell(b, 9).Value = CustDelCode;

                    //Setting the horizontal alignment to left on the fields with decimal number in there so it looks a bit better.
                    wsTom.Cell(b, 1).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsTom.Cell(b, 3).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsTom.Cell(b, 4).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsTom.Cell(b, 6).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsTom.Cell(b, 8).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
                    wsTom.Cell(b, 9).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    

                    //Raising a so that all the lines will actually be inserted
                    b++;


                }//End of while loop
                
                //Setting a auto-filter on top of the excel list
                wsToday.RangeUsed().SetAutoFilter();

                //Freezing top header
                wsToday.SheetView.Freeze(1, 9);
                //Allowing the header to be repeated on each new page
                wsToday.PageSetup.SetRowsToRepeatAtTop(1, 1);

                //Adjusting the colom widths to auto-fit 
                wsToday.Column(1).AdjustToContents();
                wsToday.Column(2).AdjustToContents();
                wsToday.Column(3).AdjustToContents();
                wsToday.Column(4).AdjustToContents();
                wsToday.Column(5).AdjustToContents();
                wsToday.Column(6).AdjustToContents();
                wsToday.Column(7).AdjustToContents();
                wsToday.Column(8).AdjustToContents();
                wsToday.Column(9).AdjustToContents();

                //setting the margins to pretty much as small as possible
                wsToday.PageSetup.Margins.Top = 0.3;
                wsToday.PageSetup.Margins.Bottom = 0.3;
                wsToday.PageSetup.Margins.Left = 0.3;
                wsToday.PageSetup.Margins.Right = 0.3;

                //Making sure it all fits onto a page
                wsToday.PageSetup.FitToPages(1, 0);

                //Setting a auto-filter on top of the excel list
                wsTom.RangeUsed().SetAutoFilter();

                //Freezing top header
                wsTom.SheetView.Freeze(1, 9);
                //Allowing the header to be repeated on each new page
                wsTom.PageSetup.SetRowsToRepeatAtTop(1, 1);

                //Adjusting the colom widths to auto-fit 
                wsTom.Column(1).AdjustToContents();
                wsTom.Column(2).AdjustToContents();
                wsTom.Column(3).AdjustToContents();
                wsTom.Column(4).AdjustToContents();
                wsTom.Column(5).AdjustToContents();
                wsTom.Column(6).AdjustToContents();
                wsTom.Column(7).AdjustToContents();
                wsTom.Column(8).AdjustToContents();
                wsTom.Column(9).AdjustToContents();

                //setting the margins to pretty much as small as possible
                wsTom.PageSetup.Margins.Top = 0.3;
                wsTom.PageSetup.Margins.Bottom = 0.3;
                wsTom.PageSetup.Margins.Left = 0.3;
                wsTom.PageSetup.Margins.Right = 0.3;

                //Making sure it all fits onto a page
                wsTom.PageSetup.FitToPages(1, 0);

                //This will be the general directory to save the files in.
                string save_dir = "C:/ProdPlan/ProductiePlanning{0}.xlsx";
                //Setting the file name here + completing the save location
                string save_loc = string.Format(save_dir, DateTime.Now.ToString("dd-MM-yyyy"));
                wb.SaveAs(save_loc);
                Console.WriteLine(save_loc);

                //ToDo:
                //Build a e-mail section which will e-mail to a distribution group and obviously attach something to it!
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //Closing the connection
            conn.Close();
            Console.WriteLine("Press enter to exit");
            Console.ReadLine();
        }
    }
}
