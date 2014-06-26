using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ZKJC.CommonDAO;

namespace TestModelTTT
{
    [Serializable]
    [TableMapping(TableName = "App_Dict")]
    public class ModelA
    {
        public string DictID { get; set; }
        public string AppID { get; set; }
        public string DictCode { get; set; }
        public string DictName { get; set; }
        [ColumnMapping(ColumnName = "ShowMode")]
        public string ShowModel { get; set; }
        public bool IsSys { get; set; }
        public string Remark { get; set; }
        public bool IsClassification { get; set; }
        public bool IsEnable { get; set; }
    }
}
