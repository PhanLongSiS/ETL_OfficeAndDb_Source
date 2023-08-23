using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransformExJsToWordJs.ModelWordJson
{
    public class WorkSpaceType
    {
        public string erroMess { get; set; } = null;
        public WorkspaceConfig config;
        /// <summary> Nội dung của các documents cần đẩy vào file word </summary>
        public DocumentsType documents;
        public WorkSpaceType()
        {
            documents = new DocumentsType();
            config = new WorkspaceConfig();
        }
    }
}
