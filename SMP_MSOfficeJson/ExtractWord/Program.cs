using ExtractWord.Models;
using Jokedst.GetOpt;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractWord
{
    public class Program
    {
        static WorkSpaceType DataInMyWord;
        /// <summary>
        /// Đọc dữ liệu từ file word và lưu vào file json
        /// </summary>
        /// <param name="args"> Tham số truyền vào command line,Ví dụ:-f vanban.docx -j abc.json</param>
        /// <returns></returns>
        static int Main(string[] args)
        {
            // dường dẫn file word cần đọc
            string path = null;
            // đường dẫn cần lưu file json
            string filejson = null;
            // bóc tách chuỗi tham số truyền vào, ví dụ: -f vanban.docx -j abc.json thì cần lấy đc văn bản cần đọc là vanban.docx và lưu vào file abc.json
            var opts = new GetOpt("Long viet cai nay", new[]
        {
            new CommandLineOption('j', "jsonoutput", "Name of the output json file",
                ParameterType.String, nextparam => filejson = (string)nextparam),
            new CommandLineOption('f', "word", "Microsoft word",
                ParameterType.String, o => path = (string)o),
        });
            opts.ParseOptions(args);
            //Đường dẫn tuyệt đối của file cần đọc. P/S: file word cần đọc phải nằm trong thư mục của AppDomain.CurrentDomain.BaseDirectory, nếu file word ko nằm trong đấy thì ko biết phải đọc file word nào
            path = Directory.GetCurrentDirectory()+"\\"+path;
            //Đường dẫn tuyệt đối của file cần đọc:P/S:file json đc tạo ra nằm trong thư mục của AppDomain.CurrentDomain.BaseDirectory
            // kiểm tra định dạng file, nếu sai định dạng trả về -1
            if (string.IsNullOrEmpty(path) || string.IsNullOrEmpty(filejson))
            {
                Console.WriteLine("Nhập sai định dạng");
                Console.ReadLine();
                return 1;
            }
            //----------------------------------- ORM To WORKBOOK  ----------------------------
            string JsonText = File.ReadAllText(filejson);
            WorkSpaceType DataOutJson;
            try
            {
                DataOutJson = JsonConvert.DeserializeObject<WorkSpaceType>(JsonText);
                if (DataOutJson==null)
                {
                    CreateFormExample(filejson);
                    return 1;
                }
            }
            catch (Exception ex)
            {
                CreateFormExample(filejson);
                return 1;
            }
            Application MyApp = new Application();
            Document MyDoc = null;
            try
            {
                if (MyApp == null)
                {
                    DataOutJson.erroMess = "Word không khởi động được.";
                    goto _END_;
                }

                // Mở file word theo đường dẫn đã cho
                MyDoc = MyApp.Documents.Open(path, ReadOnly: true);
                if (MyDoc == null)
                {
                    DataOutJson.erroMess = "Không mở được Workbook do phiên bản Word không phù hợp.";
                    return 1;
                }
            }
            catch (Exception e)
            {
                DataOutJson.erroMess = e.Message;
                return 1;
            }

            // Yêu cầu ứng dụng excel không hiển thị ra màn hình, --> chạy ngầm. //
            /* 
             * Nếu file excel DB đã được mở, không hiển thị file nữa
             * và đóng tiến trình chạy ngầm sau khi load xong dữ liệu
             */
            MyApp.Visible = DataOutJson.config.visible;

            //Set animation status for word application  
            MyApp.ShowAnimation = false;
            //----------------------------------- ORM To SHEETS  ----------------------------

            ReadFileWord(MyDoc);
            if (DataOutJson.documents.contentcontrols.Count==0 && DataOutJson.documents.tables.Count==0)
            {
                DataOutJson.erroMess="Không tồn tại bất kì ContentControl nào trong file docx";
                return 1;
            }
            foreach (var item in DataOutJson.documents.contentcontrols)
            {
                var valuetext = DataInMyWord.documents.contentcontrols.FirstOrDefault(x => x.title.Equals(item.title));
                if (valuetext != null)
                {
                    item.value = valuetext.value;
                    item.erroMess=null;
                }
                else
                {
                    item.erroMess=$"Không tồn tại ContentControl có title là {item.title} trong Word";
                    item.value=null;
                }
            }
            foreach (var item in DataOutJson.documents.tables)
            {
                var rowvalues = DataInMyWord.documents.tables.FirstOrDefault(x => x.name.Equals(item.name));
                if (rowvalues != null)
                {
                    item.rowvalues = rowvalues.rowvalues;
                    item.erroMess=null;
                }
                else
                {
                    item.erroMess=$"Không tồn tại ContentControl Table có title {item.name} này trong Word";
                    item.rowvalues=null;
                }
            }

            //----------------------------------- PRINT / SAVE  ----------------------------
            //Biến đổi thông tin đọc được thành chuỗi json
            string json = JsonConvert.SerializeObject(DataOutJson, Formatting.Indented);
            //Lưu ghi kết quả đọc đc vào file json
            try
            {

                File.WriteAllText(Directory.GetCurrentDirectory()+"\\"+filejson, json);

            }
            catch (Exception e)
            {
                DataOutJson.erroMess += (e.Message + "\n");
            }
            /// Thực hiện in luôn ra máy in tất cả các sheet nếu có yêu cầu
            if (DataOutJson.config.printnow)
            {
                /// Mặc định là máy in mặc định của hệ thống
                var printer = Type.Missing;

                /// Trường hợp có chỉ định tên máy in, thì sẽ tìm tên 
                if (DataOutJson.config.printername != string.Empty)
                {
                    var printers = PrinterSettings.InstalledPrinters;

                    foreach (String s in printers)
                    {
                        if (s.IndexOf(DataOutJson.config.printername, StringComparison.OrdinalIgnoreCase)>=0)
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
                    DataInMyWord.erroMess += (e.Message + "\n");
                }
            }

        _END_:


            // Đóng file, kết thúc
            if (DataOutJson.config.terminate)
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
            Console.WriteLine(DataOutJson.erroMess);
            foreach (var MySheetData in DataOutJson.documents.tables)
            {
                Console.WriteLine(MySheetData.erroMess);
            }
            foreach (var MySheetData in DataOutJson.documents.contentcontrols)
            {
                Console.WriteLine(MySheetData.erroMess);
            }

            //Mở file json
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = Directory.GetCurrentDirectory()+"\\"+filejson;
            prc.Start();
            return 0;
        }

        static void ReadFileWord(Document document)
        {
            try
            {
                ContentControls contentControlsCollection = document.ContentControls;//Lấy danh sách listcontententcontrol trong fileword
                DataInMyWord = new WorkSpaceType();
                List<string> list = new List<string>();
                for (int i = 1; i<=contentControlsCollection.Count; i++)
                {
                    string abc = contentControlsCollection[i].Title ?? Guid.NewGuid().ToString();
                    list.Add(abc);
                }
                for (int i = 1; i<=contentControlsCollection.Count; i++)// duyệt từng contentControl trong listcontentcontrol
                {
                    try
                    {
                        // Đọc và lưu content control thuộc bảng
                        int rows = contentControlsCollection[i].RepeatingSectionItems.Count;
                        int cells = contentControlsCollection[i].RepeatingSectionItems.Parent.Range.Cells.Count;
                        TableType tables = new TableType() { name=contentControlsCollection[i].Title.ToString() };
                        string abc = contentControlsCollection[i].Range.Text.ToString();
                        int column = cells/rows;
                        for (int j = i+1; j<=cells+i; j++)
                        {
                            RowValue row = new RowValue();
                            int index = 0;
                            while (index<column)
                            {
                                ContentControlType obj = new ContentControlType()
                                {
                                    title=contentControlsCollection[j].Title.ToString(),
                                    value=contentControlsCollection[j].Range.Text.ToString()
                                };

                                row.rowvalue.Add(obj);
                                if (index<column-1) j++;
                                index++;
                            }
                            tables.rowvalues.Add(row);
                        }
                        DataInMyWord.documents.tables.Add(tables);
                        i=i+cells;
                    }
                    catch
                    {
                        //Đọc và lưu contentcontrol ko thuộc bảng
                        if (!String.IsNullOrEmpty(contentControlsCollection[i].Title))
                        {
                            ContentControlType contentControlObject = new ContentControlType()
                            {
                                value=contentControlsCollection[i].Range.Text.ToString(),
                                title=contentControlsCollection[i].Title.ToString(),
                            };
                            DataInMyWord.documents.contentcontrols.Add(contentControlObject);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        static void CreateFormExample(string filejson)
        {
            Console.WriteLine("File mẫu không đúng định dạng, đã sinh ra form mẫu");
            List<ContentControlType> contentExample = new List<ContentControlType>()
                {
                    new ContentControlType(){title="Tên Title 1"},
                    new ContentControlType(){title="Tên Title 2"},
                    new ContentControlType(){title="Tên Title 3"},
                };
            List<TableType> tableTypesExample = new List<TableType>()
                {
                    new TableType(){name="Tên bảng 1"},
                    new TableType(){name="Tên bảng 2"},
                };
            DocumentsType documentsexaple = new DocumentsType()
            {
                contentcontrols=contentExample,
                tables=tableTypesExample,
            };
            WorkSpaceType workSpaceTypeExample = new WorkSpaceType() { config=new WorkspaceConfig(), documents=documentsexaple };
            string jsonexample = JsonConvert.SerializeObject(workSpaceTypeExample, Formatting.Indented);
            //Lưu ghi kết quả đọc đc vào file json
            File.WriteAllText(Directory.GetCurrentDirectory()+"\\"+filejson, jsonexample);
            System.Diagnostics.Process prcexaple = new System.Diagnostics.Process();
            prcexaple.StartInfo.FileName = Directory.GetCurrentDirectory()+"\\"+filejson;
            prcexaple.Start();
        }
    }
}
