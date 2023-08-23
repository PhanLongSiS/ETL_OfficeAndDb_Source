using ModifyWord.Models;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Jokedst.GetOpt;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json;

#if !DynamicLinkToTheInterop
using MSWord = Microsoft.Office.Interop.Word;       //Khong su dung nua, do da lien ket động
using System.Runtime.InteropServices;
#endif

namespace ModifyWord
{
    class Program
    {
        static public bool IsWordSupport()
        {
            Type officeType = Type.GetTypeFromProgID("Word.Application");
            return (officeType != null);
        }

        static void Main(string[] args)
        {
            //đường dẫn file Excel
            string WorldTemplateFileName = null;
            // đường dẫnfile json cần đọc
            string JsonFile = null;
            // bóc tách chuỗi tham số truyền vào, ví dụ: -f vanban.docx -j abc.json thì cần lấy đc văn bản cần đọc là vanban.docx và lưu vào file abc.json
            var opts = new GetOpt("Long viet cai nay", new[]
        {
            new CommandLineOption('j', "jsoninput", "Name of the output json file",
                ParameterType.String, nextparam => JsonFile = (string)nextparam),
            new CommandLineOption('f', "excel", "Microsoft Excel",
                ParameterType.String, o => WorldTemplateFileName = (string)o),
        });
            opts.ParseOptions(args);
            // kiểm tra định dạng file, nếu sai định dạng trả về -1
            if (string.IsNullOrEmpty(WorldTemplateFileName) || string.IsNullOrEmpty(JsonFile))
            {
                Console.WriteLine("Nhập sai định dạng");
                Console.ReadLine();
                return ;
            }
            string errMsg = null;

            if (!IsWordSupport())
            {
                errMsg = "Chưa cài đặt MS Word.";
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

            WorldTemplateFileName = Directory.GetCurrentDirectory() + "\\" + WorldTemplateFileName;
            if (!File.Exists(WorldTemplateFileName))
            {
                errMsg = "File " + WorldTemplateFileName + " không tôn tại";
                json = "";

            }

            Modification2(json, WorldTemplateFileName);

        _END_:
            Console.WriteLine("Version 2.0");
            Console.WriteLine(errMsg);
            Debug.WriteLine(errMsg);
        }

        /// <summary>
        ///     Nhồi dữ liệu vào file excel
        /// </summary>
        /// <param name="jsontext"> Chứa thông tin cần nhồi vào file excel </param>
        /// <param name="wordTemplateFileName" File Excel template gốc </param> 
        private static void Modification2(string jsontext, string wordTemplateFileName)
        {
            //-----------------------------------JSON PARSER to  ORM ----------------------------
            /// Đọc và phân tích nội dung json đầu vào
            JsonTextReader reader = new JsonTextReader(new StringReader(jsontext));


            WorkspaceType MyDocumentStructure = null;

            /// Nếu dữ liệu đầu vào không có, thì sẽ hiển thị cấu trúc dữ liệu demo 
            if (jsontext == "")
            {
                /// Đối tượng json handler
                JsonSerializer serializer = new JsonSerializer();

                MyDocumentStructure = new WorkspaceType();
                WorkspaceConfig config=new WorkspaceConfig();

                config.printername = "pdf";
                 config.visible = true;
                config.saveas = "demo.docx";
                ///    - Bổ sung cấu hình
                MyDocumentStructure.config=config;

                ///    - Bổ sung sheet mới để có dữ liệu minh họa
                DocumentType documentelements = new DocumentType(new List<ContentControlType>());
                MyDocumentStructure.documents=documentelements;

                ///    - Bổ sung các cell mới vào sheet nói trên, để có dữ liệu minh họa
                documentelements.contentcontrols.Add(new ContentControlType("title", true));
                documentelements.contentcontrols.Add(new ContentControlType("name", "Text"));
                documentelements.contentcontrols.Add(new ContentControlType("date", "20/11/2012"));

                ////    - Ghi file demo
                using (StreamWriter sw = new StreamWriter(@"demo.json"))
                using (JsonWriter writer = new JsonTextWriter(sw))
                {
                    serializer.Serialize(writer, MyDocumentStructure);
                }
                return;
            }

            try
            {
                //ORM hóa nội dung json vào class
                MyDocumentStructure = JsonConvert.DeserializeObject<WorkspaceType>(jsontext);
            }
            catch (Exception e)
            {
                Console.WriteLine("Json Error: " + e.Message);
                Debug.WriteLine("Json Error: " + e.Message);
                return;
            }

            //----------------------------------- ORM To WORKBOOK  ----------------------------


            Microsoft.Office.Interop.Word.Application MyApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document MyDoc = null;
            try
            {
                if (MyApp == null)
                {
                    MyDocumentStructure.errMessage = "Word không khởi động được.";
                    goto _END_;
                }

                // Mở file word theo đường dẫn đã cho
                MyDoc = MyApp.Documents.Open(wordTemplateFileName, ReadOnly: true);
                if (MyDoc == null)
                {
                    MyDocumentStructure.errMessage = "Không mở được Workbook do phiên bản Word không phù hợp.";
                    goto _END_;
                }
            }
            catch (Exception e)
            {
                MyDocumentStructure.errMessage = e.Message;
                goto _END_;
            }

            // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
            /* 
             * Nếu file excel DB đã được mở, không hiển thị file nữa
             * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
             */
            MyApp.Visible = MyDocumentStructure.config.visible;

            //Set animation status for word application  
            MyApp.ShowAnimation = false;

            //----------------------------------- ORM To SHEETS  ----------------------------
            DocumentType MyDocumentData = MyDocumentStructure.documents;
            {
                // Xác định tất cả các control content đang có trong file
                // The code below search content controls in all word document stories see http://word.mvps.org/faqs/customization/ReplaceAnywhere.htm
                List<MSWord.ContentControl> ccList = new List<MSWord.ContentControl>();

                MSWord.Range rangeStory;
                foreach (MSWord.Range range in MyDoc.StoryRanges)
                {
                    rangeStory = range;
                    do
                    {
                        try
                        {
                            foreach (MSWord.ContentControl cc in range.ContentControls)
                            {
                                ccList.Add(cc);
                            }

                            // Get the content controls in the shapes ranges
                            foreach (MSWord.Shape shape in range.ShapeRange)
                            {
                                foreach (MSWord.ContentControl cc in shape.TextFrame.TextRange.ContentControls)
                                {
                                    ccList.Add(cc);
                                }

                            }
                        }
                        catch (COMException) { }
                        rangeStory = rangeStory.NextStoryRange;

                    }
                    while (rangeStory != null);
                }

                //Điền nội dung vào các content control
                foreach (MSWord.ContentControl cc in ccList)
                {
                    ContentControlType myData = MyDocumentData.contentcontrols.Find(item => item.title == cc.Title);
                    if (myData != null)
                    {
                        //cc.SetPlaceholderText(null, null, myData.value.ToString());
                        cc.Range.Text = myData.value.ToString();
                    }
                }
            }

            //----------------------------------- PRINT / SAVE  ----------------------------

            /// Lưu lại file đã có kết quả
            if (MyDocumentStructure.config.saveas != string.Empty)
            {
                if (!System.IO.Path.IsPathRooted(MyDocumentStructure.config.saveas))
                {// Đường dẫn mặc định là thư mục Documents. Cần đổi về đường dẫn tương đối
                    MyDocumentStructure.config.saveas = Directory.GetCurrentDirectory() + @"\" + MyDocumentStructure.config.saveas;
                }
                try
                {
                    MyApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone;

                    //Lưu ý: Do muốn overwritten, mà hằng Excel.XlSaveAsAccessMode.xlNoChange lại không có khi triệu gọi lb động --> sử dụng luôn hằng số 1.
                    MyDoc.SaveAs(MyDocumentStructure.config.saveas, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, 1,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MyApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsAll;
                }
                catch (Exception e)
                {
                    MyDocumentStructure.errMessage += (e.Message + "\n");
                    Debug.WriteLine(e.Message);
                    goto _END_;
                }
            }

            /// Thực hiện in luôn ra máy in tất cả các sheet nếu có yêu cầu
            if (MyDocumentStructure.config.printnow)
            {
                /// Mặc định là máy in mặc định của hệ thống
                var printer = Type.Missing;

                /// Trường hợp có chỉ định tên máy in, thì sẽ tìm tên 
                if (MyDocumentStructure.config.printername != string.Empty)
                {
                    var printers = PrinterSettings.InstalledPrinters;

                    foreach (String s in printers)
                    {
                        if (s.IndexOf(MyDocumentStructure.config.printername, StringComparison.OrdinalIgnoreCase)>=0)
                        {
                            printer = s;
                            break;
                        }
                    }
                }
                try
                {
                    MyDoc.PrintOut(
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        printer, Type.Missing, Type.Missing, Type.Missing);
                }
                catch (Exception e)
                {
                    MyDocumentStructure.errMessage += (e.Message + "\n");
                }
            }

        _END_:


            // Đóng file, kết thúc
            if (MyDocumentStructure.config.terminate)
            {
                if (MyDoc != null)
                {
                    MyDoc.Close(SaveChanges: false);
                }
                if (MyApp != null)
                {
                    MyApp.Quit();
                }
            }
            Console.WriteLine("Version 2.0");

            /// Hiển thị tất cả các lỗi
            Console.WriteLine(MyDocumentStructure.errMessage);
        }
    }
}
