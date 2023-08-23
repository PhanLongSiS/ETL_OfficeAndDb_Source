using Jokedst.GetOpt;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TransformWordJsToExJs.ModelExcelJson;
using TransformWordJsToExJs.ModelWordJson;

namespace TransformWordJsToExJs
{
    public class Program
    {
        static int Main(string[] args)
        {
            // dường dẫn file word cần đọc
            string WordFile = null;
            // bóc tách chuỗi tham số truyền vào, ví dụ: -f vanban.docx -j abc.json thì cần lấy đc văn bản cần đọc là vanban.docx và lưu vào file abc.json
            var opts = new GetOpt("Long viet cai nay", new[]
        {
            new CommandLineOption('w', "excel", "Name of the output json file",
                ParameterType.String, nextparam => WordFile = (string)nextparam),
        });
            opts.ParseOptions(args);
            if (string.IsNullOrEmpty(WordFile))
            {
                Console.WriteLine("Nhập sai định dạng");
                return 0;
            }
            string json;
            string errMsg = null;
            if (!File.Exists(WordFile))
            {
                errMsg = "File " + WordFile + " không tôn tại";
                json = "";

            }
            else
            {
                json = File.ReadAllText(WordFile);
            }
            TransformJson(json, WordFile);
            return 1;
        }
        static void TransformJson(string JsonText, string WordJsonFile)
        {
            WorkSpaceType MyWordSpaceData = new WorkSpaceType();
            try
            {
                /* ORM hóa nội dung json vào class "WorkbookType"
                 * Nếu ko thể ORM thì thông báo nhập sai cấu trúc của file json, sau đó tạo bản ghi mẫu cho người dùng
                 */
                MyWordSpaceData = JsonConvert.DeserializeObject<WorkSpaceType>(JsonText);
                var MyWorkBookData = new WorkbookType();
                //Tranform Config
                MyWorkBookData.config.terminate=MyWordSpaceData.config.terminate;
                MyWorkBookData.config.visible=MyWordSpaceData.config.visible;
                MyWorkBookData.config.printnow=MyWordSpaceData.config.printnow;
                MyWorkBookData.config.activatesheet=true;
                MyWorkBookData.config.printername=MyWordSpaceData.config.printername;
                //Get SheetName
                SheetType sheet=new SheetType();
               var WordJsonFileTD = Directory.GetCurrentDirectory() + "\\" + WordJsonFile;
               var SheetName=Path.GetFileNameWithoutExtension(WordJsonFileTD);
                sheet.name=SheetName;
                //TranformData
                foreach (var wordspace in MyWordSpaceData.documents.contentcontrols)
                {
                    sheet.cells.Add(new CellType() { posName=wordspace.title, value=wordspace.value });
                }
                MyWorkBookData.sheets.Add(sheet);
                string jsonExcel = JsonConvert.SerializeObject(MyWorkBookData, Formatting.Indented);
                File.WriteAllText(Directory.GetCurrentDirectory()+"\\"+WordJsonFile, jsonExcel);
            }
            catch (Exception)
            {
                Console.WriteLine("Không đúng định dạng json Excel");
                return;
            }
        }
    }
}
