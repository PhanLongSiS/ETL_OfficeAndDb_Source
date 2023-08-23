using ExtractExcel.Models;
using Jokedst.GetOpt;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExtractExcel
{
    public class Program
    {
        /// <summary>
        /// Kiểm tra trên máy tính đã cài đặt phần mềm Excel hay chưa
        /// </summary>
        /// <returns>True: Đã cài đăt</returns>
        /// <returns>False: Chưa cài đăt</returns>
        static public bool IsExcelSupport()
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            return (officeType != null);
        }
        static int Main(string[] args)
        {
            //đường dẫn file Excel cần đọc
            string ExcelTemplateFileName = null;
            // đường dẫnfile json cần ghi
            string JsonFile = null;
            // bóc tách chuỗi tham số truyền vào, ví dụ: -f vanban.docx -j abc.json thì cần lấy đc văn bản cần đọc là vanban.docx và lưu vào file abc.json
            var opts = new GetOpt("Long viet cai nay", new[]
        {
            new CommandLineOption('j', "jsoninput", "Name of the output json file",
                ParameterType.String, nextparam => JsonFile = (string)nextparam),
            new CommandLineOption('f', "excel", "Microsoft Excel",
                ParameterType.String, o => ExcelTemplateFileName = (string)o),
        });
            opts.ParseOptions(args);
            // kiểm tra định dạng file, nếu sai định dạng trả về -1
            if (string.IsNullOrEmpty(ExcelTemplateFileName) || string.IsNullOrEmpty(JsonFile))
            {
                Console.WriteLine("Nhập sai định dạng");
                Console.ReadLine();
                return 1;
            }
            string errMsg = null;

            if (!IsExcelSupport())
            {
                errMsg = "Chưa cài đặt MS Excel.";
                return 1;
            }
            string json;

            if (!File.Exists(JsonFile))
            {
                errMsg = "File " + JsonFile + " không tôn tại";
                json = "";
                Console.WriteLine(errMsg);
                return 1;
            }
            else
            {
                json = File.ReadAllText(JsonFile);
            }

            ExcelTemplateFileName = Directory.GetCurrentDirectory() + "\\" + ExcelTemplateFileName;
            if (!File.Exists(ExcelTemplateFileName))
            {
                errMsg = "File " + ExcelTemplateFileName + " không tôn tại ";
                json = "";
                Console.WriteLine(errMsg);
                return 1;

            }
                 ExtractExcel2(JsonFile, ExcelTemplateFileName);
                System.Diagnostics.Process prc = new System.Diagnostics.Process();
                prc.StartInfo.FileName = Directory.GetCurrentDirectory()+"\\"+JsonFile;
                prc.Start();
                Console.WriteLine("Version 2.0");
                return 0;
        }
        /// <summary>
        ///     Lấy dữ liệu từ file excel -> json
        /// </summary>
        /// <param name="jsontext"> Chứa thông tin cần nhồi vào file excel </param>
        /// <param name="excelTemplateFileName" File Excel template gốc </param> 
        private static void ExtractExcel2(string jsonfile, string excelpath)
        {
            string JsonText = File.ReadAllText(jsonfile);
            WorkbookType MyWorkbookData = new WorkbookType();
            try
            {
                /* ORM hóa nội dung json vào class "WorkbookType"
                 * Nếu ko thể ORM thì thông báo nhập sai cấu trúc của file json, sau đó tạo bản ghi mẫu cho người dùng
                 */
                MyWorkbookData = JsonConvert.DeserializeObject<WorkbookType>(JsonText);
            }
            catch (Exception)
            {
                CreateExample(jsonfile);
                return ;
            }
            if(MyWorkbookData == null)
            {
                CreateExample(jsonfile);
                return ;
                
            }
            //Tạo các biến thuộc kiểu dữ liệu ở thư viện Interop Excel để sử dụng thư viện Interop Excel
            Excel.Application xlApp;
            Excel.Workbook xlWordBook;
            Excel.Worksheet xlworkSheet = null;
            Excel.Range range;
            try
            {
                xlApp =new Excel.Application();
                if (xlApp == null)
                {
                    MyWorkbookData.errMessage = "Excel không khởi động được.";
                    return ;
                }

                // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
                /* 
                 * Nếu file excel DB đã được mở, không hiển thị file nữa
                 * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
                 */
                xlApp.Visible = MyWorkbookData.config.visible;

                // Mở file excel theo đường dẫn đã cho
                xlWordBook = xlApp.Workbooks.Open(excelpath, ReadOnly: true);
            }
            catch (Exception e)
            {

                MyWorkbookData.errMessage = e.Message;
                return;
            }
            WorkbookType newWorkbook = new WorkbookType();
            /*
             * Bước 1: 
             */
            foreach (var mysheetOld in MyWorkbookData.sheets)//Duyệt lần lượt sheet có trong file json(trang sheet cần đọc)
            {
                for (int i = 1; i<=xlWordBook.Worksheets.Count; i++)//Duyệt lần lượt sheet có trong file Excel
                {
                    /* checkSheetExit kiểm tra trang sheet trong file Json có tồn tại trong file Excel hay không
                     * True:Có tồn tại trong file Excel
                     * False:Không tồn tại trong file Excel
                     */
                    bool checkSheetExit = false;

                    //Lấy ra sheet
                    xlworkSheet=(Excel.Worksheet)xlWordBook.Worksheets.get_Item(i);

                    //Lấy ra tên sheet 
                    string name = xlworkSheet.Name;
                    range=xlworkSheet.UsedRange;

                    //Kiểm tra sheet đang được duyệt đến có tên trùng với tên sheet đang cần đọc
                    if (name.Equals(mysheetOld.name))
                    {
                        SheetType mySheetNew = new SheetType() { name=name };
                        foreach (var cellold in mysheetOld.cells)// Đọc gái trị các cell
                        {
                            if (!string.IsNullOrEmpty(cellold.pos))
                            {
                                CellPosition cellPosition = new CellPosition(cellold.pos);
                                var value = Convert.ToString((range.Cells[cellPosition.RowIndex, cellPosition.ColumnIndex] as Excel.Range).Value2);
                                try
                                {
                                    var posname = ((range.Cells[cellPosition.RowIndex, cellPosition.ColumnIndex] as Excel.Range).Name).Name;
                                    CellType newcell = new CellType() { pos=cellold.pos, value=value, posName=posname };
                                    mySheetNew.cells.Add(newcell);
                                }
                                catch (Exception)
                                {
                                    CellType newcell = new CellType() { pos=cellold.pos, value=value, posName=null };
                                    mySheetNew.cells.Add(newcell);
                                }
                            }
                            else
                            {
                                var cell = xlworkSheet.get_Range(cellold.posName, cellold.posName);
                                var value = Convert.ToString((range.Cells[cell.Row, cell.Column] as Excel.Range).Value2);
                                CellType newcell = new CellType() { pos="", value=value, posName=cellold.posName };
                                mySheetNew.cells.Add(newcell);
                            }
                        }
                        newWorkbook.sheets.Add(mySheetNew);
                        checkSheetExit=true;
                        break;
                    }

                    //Nếu đã duyệt hết tất cả các sheet trong file Excel mà vẫn ko tìm thấy sheet đang cần đọc=>Sheet không tồn tại
                    if (i==xlWordBook.Worksheets.Count && !checkSheetExit)
                    {
                        SheetType mysheetErrol = new SheetType() { name=mysheetOld.name, errMessage=$"Trang sheet co ten {mysheetOld.name} khong ton tai trong file Excel nay", cells=new List<CellType>() };
                        newWorkbook.sheets.Add(mysheetErrol);
                    }
                }
            }
            //Convert đối đượng newWorkbook thành chuỗi json
            string jsonSeri = JsonConvert.SerializeObject(newWorkbook, Formatting.Indented);
            //Lưu ghi kết quả đọc đc vào file json
            try
            {

                File.WriteAllText(Directory.GetCurrentDirectory()+"\\"+jsonfile, jsonSeri);

            }
            catch (Exception e)
            {
                newWorkbook.errMessage += (e.Message + "\n");
                return ;
            }
            //----------------------------------- PRINT / SAVE  ----------------------------

            /// Thực hiện in luôn ra máy in tất cả các sheet nếu có yêu cầu
            if (MyWorkbookData.config.printnow)
            {
                /// Mặc định là máy in mặc định của hệ thống
                var printer = Type.Missing;

                /// Trường hợp có chỉ định tên máy in, thì sẽ tìm tên 
                if (MyWorkbookData.config.printername != string.Empty)
                {
                    var printers = PrinterSettings.InstalledPrinters;

                    foreach (String s in printers)
                    {
                        if (s.IndexOf(MyWorkbookData.config.printername, StringComparison.OrdinalIgnoreCase)>=0)
                        {
                            printer = s;
                            break;
                        }
                    }
                }

                foreach (SheetType MySheetData in MyWorkbookData.sheets)
                {
                    //Mở lại sheet đó
                    if (MySheetData.name != null)
                    {
                        if (Int32.TryParse(MySheetData.name, out int SheetIndex))
                        {
                            xlworkSheet = xlWordBook.Worksheets[SheetIndex];
                        }
                        else
                        {
                            xlworkSheet = xlWordBook.Worksheets[MySheetData.name];
                        }
                    }
                    //và in đúng máy in chỉ định
                    try
                    {
                        xlworkSheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            printer, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch (Exception e)
                    {
                        newWorkbook.errMessage += (e.Message + "\n");
                        continue;
                    }
                }
            }

        _END_:


            // Đóng file, kết thúc
            if (MyWorkbookData.config.terminate)
            {
                if (xlWordBook != null)
                {
                    xlWordBook.Close(SaveChanges: false);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                }
            }
            Console.WriteLine("Version 2.0");

            /// Hiển thị tất cả các lỗi
            Console.WriteLine(newWorkbook.errMessage);
            foreach (SheetType MySheetData in newWorkbook.sheets)
            {
                Console.WriteLine(MySheetData.errMessage);
            }
        }
        static void CreateExample(string jsonfile)
        {
            Console.WriteLine(@"Du lieu file json khong dung dinh dang, da sinh ra file mau");
            //Sinh ra một cấu trúc file json mẫu cho người dùng
            WorkbookType example = new WorkbookType();
            SheetType exampleSheet1 = new SheetType()
            {
                name="Tên sheet(VD:Hợp đồng vật tư y tế)",
                cells=new List<CellType>()
                    {
                        new CellType(){pos="A5",value=""},
                        new CellType(){pos="B12",value=""},
                        new CellType(){pos="A5",value=""}
                    }
            };
            SheetType exampleSheet2 = new SheetType()
            {
                name="Tên sheet(VD:Hợp đồng vật tư y tế)",
                cells=new List<CellType>()
                    {
                        new CellType(){pos="A5",value=""},
                        new CellType(){pos="A9",value=""},
                        new CellType(){pos="C5",value=""}
                    }
            };
            example.sheets.Add(exampleSheet1);
            example.sheets.Add(exampleSheet2);
            string jsonSeriexample = JsonConvert.SerializeObject(example, Formatting.Indented);
            File.WriteAllText(jsonfile, jsonSeriexample);
        }

    }
}
