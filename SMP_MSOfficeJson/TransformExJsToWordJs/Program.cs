using Jokedst.GetOpt;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TransformExJsToWordJs.ModelExcelJson;
using TransformExJsToWordJs.ModelWordJson;

namespace TransformExJsToWordJs
{
    public class Program
    {
        static int Main(string[] args)
        {
            // dường dẫn file word cần đọc
            string ExcelFile = null;
            // đường dẫn cần lưu file json
            // bóc tách chuỗi tham số truyền vào, ví dụ: -f vanban.docx -j abc.json thì cần lấy đc văn bản cần đọc là vanban.docx và lưu vào file abc.json
            var opts = new GetOpt("Long viet cai nay", new[]
        {
            new CommandLineOption('e', "excel", "Name of the output json file",
                ParameterType.String, nextparam => ExcelFile = (string)nextparam),
        });
            opts.ParseOptions(args);
            if (string.IsNullOrEmpty(ExcelFile))
            {
                Console.WriteLine("Nhập sai định dạng");
                return 0;
            }
            string json;
            string errMsg = null;
            if (!File.Exists(ExcelFile))
            {
                errMsg = "File " + ExcelFile + " không tôn tại";
                json = "";

            }
            else
            {
                json = File.ReadAllText(ExcelFile);
            }
            TransformJson(json,ExcelFile);
            return 1;
        }
        static void TransformJson(string JsonText,string ExcelJsonFile)
        {
            WorkbookType MyWorkbookData = new WorkbookType();
            try
            {
                /* ORM hóa nội dung json vào class "WorkbookType"
                 * Nếu ko thể ORM thì thông báo nhập sai cấu trúc của file json, sau đó tạo bản ghi mẫu cho người dùng
                 */
                MyWorkbookData = JsonConvert.DeserializeObject<WorkbookType>(JsonText);
                var MyWorkSpaceData = new WorkSpaceType();
                //Tranform Config
                MyWorkSpaceData.config.terminate=MyWorkbookData.config.terminate;
                MyWorkSpaceData.config.visible=MyWorkbookData.config.visible;
                MyWorkSpaceData.config.printnow=MyWorkbookData.config.printnow;
                MyWorkSpaceData.config.printername=MyWorkbookData.config.printername;
                //TranformData
                var MyDocumentsType=new DocumentsType();
                foreach(var valuecell in MyWorkbookData.sheets[0].cells)
                {
                    if (valuecell.posName!=null)
                    {
                        MyDocumentsType.contentcontrols.Add(new ContentControlType() { title=valuecell.posName, value=valuecell.value });
                    }
                }
                MyWorkSpaceData.documents=MyDocumentsType;
                string jsonWord = JsonConvert.SerializeObject(MyWorkSpaceData, Formatting.Indented);
                File.WriteAllText(Directory.GetCurrentDirectory()+"\\"+ExcelJsonFile, jsonWord);
            }
            catch (Exception)
            {
                Console.WriteLine("Không đúng định dạng json Excel");
                return;
            }
        }
    }
}
