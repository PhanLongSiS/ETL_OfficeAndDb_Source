using Jokedst.GetOpt;
using ModifyExcel.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;       //Khong su dung nua, do da lien ket động

namespace ModifyExcel
{
    class Program
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
            //đường dẫn file Excel
            string ExcelTemplateFileName = null;
            // đường dẫnfile json cần đọc
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
                goto _END_;
            }
            string json;

            if (!File.Exists(JsonFile))
            {
                errMsg = "File " + JsonFile + " không tôn tại";
                json = "";

            }
            else
            {
                json = File.ReadAllText(JsonFile);
            }

            ExcelTemplateFileName = Directory.GetCurrentDirectory() + "\\" + ExcelTemplateFileName;
            if (!File.Exists(ExcelTemplateFileName))
            {
                errMsg = "File " + ExcelTemplateFileName + " không tôn tại";
                json = "";

            }

            Modification2(json, ExcelTemplateFileName, JsonFile);

        _END_:
            Console.WriteLine("Version 2.0");
            Console.WriteLine(errMsg);
            Debug.WriteLine(errMsg);
            return 0;
        }

        /// <summary>
        ///     Nhồi dữ liệu vào file excel
        /// </summary>
        /// <param name="jsontext"> Chứa thông tin cần nhồi vào file excel </param>
        /// <param name="excelTemplateFileName" File Excel template gốc </param> 
        private static void Modification2(string jsontext, string excelTemplateFileName,string jsonfile)
        {
            //-----------------------------------JSON PARSER to  ORM ----------------------------
            /// Đọc và phân tích nội dung json đầu vào
            JsonTextReader reader = new JsonTextReader(new StringReader(jsontext));


            WorkbookType MyWorkbookData;

            /// Nếu dữ liệu đầu vào không có, thì sẽ hiển thị cấu trúc dữ liệu demo 
            if (jsontext == "")
            {
                /// Đối tượng json handler
                JsonSerializer serializer = new JsonSerializer();

                MyWorkbookData = new WorkbookType();

                ///    - Bổ sung cấu hình
                MyWorkbookData.config.printername = "pdf";
                MyWorkbookData.config.visible = true;
                MyWorkbookData.config.saveas = "demo.xlsx";

                ///    - Bổ sung sheet mới để có dữ liệu minh họa
                SheetType MySheetData = new SheetType(new List<CellType>(), "Programable Sheet");
                MyWorkbookData.sheets.Add(MySheetData);

                ///    - Bổ sung các cell mới vào sheet nói trên, để có dữ liệu minh họa
                MySheetData.cells.Add(new CellType("A5", true));
                MySheetData.cells.Add(new CellType("B1", "Text"));
                MySheetData.cells.Add(new CellType("C3", "20/11/2012"));
                MySheetData.name = "example";

                ////    - Ghi file demo
                string jsonSeriexample = JsonConvert.SerializeObject(MyWorkbookData, Formatting.Indented);
                File.WriteAllText(jsonfile, jsonSeriexample);
                return;
            } // Kết thúc phần sinh dữ liệu minh hoạ

            try
            {
                //ORM hóa nội dung json vào class
                MyWorkbookData = JsonConvert.DeserializeObject<WorkbookType>(jsontext);
            }
            catch (Exception e)
            {
                Console.WriteLine("Json Error: " + e.Message);
                Debug.WriteLine("Json Error: " + e.Message);
                return;
            }

            //----------------------------------- ORM To WORKBOOK  ----------------------------

            /// Triệu gọi thư viện Interop Excel kiểu động.
            Type typeExcel = Type.GetTypeFromProgID("Excel.Application");
            dynamic Excel = Activator.CreateInstance(typeExcel);

            string errMsg = null;

            /// Tạo các biến thuộc kiểu dữ liệu liên kết động ở thư viện  Interop Excel. Thế là xong. Mọi việc lại diễn ra bình thường
            dynamic MyBook = null;
            dynamic MyApp = null;
            Excel.Worksheet MySheet = null;
            try
            {
                MyApp = Excel.Application();
                if (MyApp == null)
                {
                    MyWorkbookData.errMessage = "Excel không khởi động được.";
                    goto _END_;
                }

                // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
                /* 
                 * Nếu file excel DB đã được mở, không hiển thị file nữa
                 * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
                 */
                MyApp.Visible = MyWorkbookData.config.visible;

                // Mở file excel theo đường dẫn đã cho
                MyBook = MyApp.Workbooks.Open(excelTemplateFileName, ReadOnly: true);
                if (MyBook == null)
                {
                    MyWorkbookData.errMessage = "Không mở được Workbook do phiên bản Excel không phù hợp.";
                    goto _END_;
                }
            }
            catch (Exception e)
            {
                MyWorkbookData.errMessage = e.Message;
                goto _END_;
            }
            //----------------------------------- ORM To SHEETS  ----------------------------
            foreach (SheetType MySheetData in MyWorkbookData.sheets)
            {
                //Mở Sheet có tên như chỉ định, hoặc sử dụng số thứ tự
                try
                {
                    if (MySheetData.name != null)
                    {
                        if (Int32.TryParse(MySheetData.name, out int SheetIndex))
                        {
                            MySheet = MyBook.Worksheets[SheetIndex];
                        }
                        else
                        {
                            MySheet = MyBook.Worksheets[MySheetData.name];
                        }

                        // Nếu có activate thì focus vào sheet này
                        if (MyWorkbookData.config.activatesheet)
                        {
                            MySheet.Activate();
                        }
                    }
                    else
                    {
                        MySheetData.errMessage += "Tên Sheet trống, không hợp lệ.\n";
                        continue;
                    }
                }
                catch
                {
                    MySheetData.errMessage = "File dữ liệu Excel không có worksheet có tên qui định là " + MySheetData.name + ". Các sheet đang có là";
                    foreach (dynamic sh in MyBook.Worksheets)
                    {
                        MySheetData.errMessage = MySheetData.errMessage + ";" + sh.Name;
                    }

                    Debug.WriteLine(MySheetData.errMessage);
                    Console.WriteLine(MySheetData.errMessage);
                    continue;
                }


                //Ẩn hoặc hiện sheet. 3 trạng thái sau chỉ áp dụng cho NewSheet: xlSheetVeryHidden thì chỉ có thể cho hiện lại bằng lập trình, xlSheetHidden thì người dùng có thể tự unhide được
                //MySheet.Visible = MySheetData.visible ? Excel.XlSheetVisibility.xlSheetVisible : Excel.XlSheetVisibility.xlSheetVeryHidden;


                /// Ghi nội dung vào các cell
                MySheetData.errMessage = "";
                foreach (CellType MyCellData in MySheetData.cells)
                {
                    try
                    {
                        if(string.IsNullOrEmpty(MyCellData.pos))
                        {
                            var cell = MySheet.get_Range(MyCellData.posname, MyCellData.posname);
                            MySheet.Cells[cell.Row, cell.Column].Value = MyCellData.value;
                        }
                        else
                        {
                            MySheet.Cells[MyCellData.RowIndex, MyCellData.ColumnIndex].Value = MyCellData.value;
                        }
                    }
                    catch (Exception e)
                    {
                        MySheetData.errMessage += (MyCellData.pos + " " + e.Message);
                        Debug.WriteLine(MySheetData.errMessage);
                        continue;
                    }
                }
            }

            //----------------------------------- PRINT / SAVE  ----------------------------

            /// Lưu lại file đã có kết quả
            if (MyWorkbookData.config.saveas != string.Empty)
            {
                if (!System.IO.Path.IsPathRooted(MyWorkbookData.config.saveas))
                {// Đường dẫn mặc định là thư mục Documents. Cần đổi về đường dẫn tương đối
                    MyWorkbookData.config.saveas = Directory.GetCurrentDirectory() + @"\" + MyWorkbookData.config.saveas;
                }
                try
                {
                    MyApp.DisplayAlerts = false;
                    //Lưu ý: Do muốn overwritten, mà hằng Excel.XlSaveAsAccessMode.xlNoChange lại không có khi triệu gọi lb động --> sử dụng luôn hằng số 1.
                    MyBook.SaveAs(MyWorkbookData.config.saveas, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MyApp.DisplayAlerts = true;
                }
                catch (Exception e)
                {
                    MyWorkbookData.errMessage += (e.Message + "\n");
                    Debug.WriteLine(e.Message);
                    goto _END_;
                }
            }

            /// Thực hiện in luôn ra máy in tất cả các sheet nếu có yêu cầu
            if (MyWorkbookData.config.printnow)
            {
                /// Mặc định là máy in mặc định của hệ thống
                var printer = Type.Missing;

                /// Trường hợp có chỉ định tên máy in, thì sẽ tìm tên 
                if (MyWorkbookData.config.printername != string.Empty)
                {
                    var printers =PrinterSettings.InstalledPrinters;

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
                            MySheet = MyBook.Worksheets[SheetIndex];
                        }
                        else
                        {
                            MySheet = MyBook.Worksheets[MySheetData.name];
                        }
                    }
                    //và in đúng máy in chỉ định
                    try
                    {
                        MySheet.PrintOut(
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            printer, Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch (Exception e)
                    {
                        MyWorkbookData.errMessage += (e.Message + "\n");
                        continue;
                    }
                }
            }

        _END_:


            // Đóng file, kết thúc
            if (MyWorkbookData.config.terminate)
            {
                if (MyBook != null)
                {
                    MyBook.Close(SaveChanges: false);
                }
                if (MyApp != null)
                {
                    MyApp.Quit();
                }
            }
            Console.WriteLine("Version 2.0");

            /// Hiển thị tất cả các lỗi
            Console.WriteLine(MyWorkbookData.errMessage);
            foreach (SheetType MySheetData in MyWorkbookData.sheets)
            {
                Console.WriteLine(MySheetData.errMessage);
            }
        }
    }
}
