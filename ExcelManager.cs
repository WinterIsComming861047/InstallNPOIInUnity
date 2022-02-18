using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace prjCGSSim
{
    public class ExcelManager
    {

        static List<Dictionary<string, object>> ExcelData = readExcel();
        public static List<Dictionary<string, object>> readExcel()
        {
            string filePath = @"C:\Users\winte\Downloads\ExcelDataForNPOI.xlsx";
            IWorkbook wk = null;
            string extension = Path.GetExtension(filePath);
            List<Dictionary<string, object>> stocklist = new List<Dictionary<string, object>>();
            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".xls"))
                {
                    //把xls檔案中的資料寫入wk中
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    //把xlsx檔案中的資料寫入wk中
                    wk = new XSSFWorkbook(fs);
                }

                fs.Close();
                //讀取當前表資料
                ISheet sheet = wk.GetSheetAt(0);
                IRow row = sheet.GetRow(0);  //讀取當前行資料



                for (int i = 1; i <= sheet.LastRowNum; i++)  //LastRowNum 是當前表的總行數-1（注意）
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    row = sheet.GetRow(i);  //讀取當前行資料
                    if (row != null)
                    {
                        //LastCellNum 是當前行的總列數
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell.CellType == CellType.Numeric)
                            {
                                //NPOI中數字和日期都是NUMERIC類型的，這裏對其進行判斷是否是日期類型
                                if (HSSFDateUtil.IsCellDateFormatted(cell))//日期類型
                                {
                                    //讀取該行的第j列資料
                                    object value = cell.DateCellValue.ToShortDateString();
                                    dict.Add(sheet.GetRow(0).GetCell(j).ToString(), value);
                                }
                                else//其他數字類型
                                {
                                    object value = cell.NumericCellValue;
                                    dict.Add(sheet.GetRow(0).GetCell(j).ToString(), value);
                                }
                            }
                            else
                            {
                                object value = cell.ToString();
                                dict.Add(sheet.GetRow(0).GetCell(j).ToString(), value);
                            }

                            //Console.Write(value.ToString() + " ");
                        }
                        stocklist.Add(dict);
                        //Console.WriteLine("\t");
                    }
                }
                //Console.ReadKey();
                return stocklist;

            }

            catch (Exception e)
            {
                //只在Debug模式下才輸出
                Console.WriteLine(e.Message);
                return stocklist;

            }

        }

        public static string OY01StorageCode(string ColumnName)
        {
            string StorageCode;
            foreach (Dictionary<string, object> data in ExcelData)
            {
                string tempArea = data["儲區"].ToString() + data["儲位"].ToString();               
                if (tempArea == ColumnName)
                {
                    StorageCode = data["代碼"].ToString();
                    return StorageCode;
                }
            }
            StorageCode = "";
            return StorageCode;
        }
        public static string OY01Layer(string ColumnName)
        {
            string Layer;
            foreach (Dictionary<string, object> data in ExcelData)
            {
                string tempArea = data["儲區"].ToString() + data["儲位"].ToString();
                if (tempArea == ColumnName)
                {
                    Layer = data["層數"].ToString();
                    return Layer;
                }
            }
            Layer = "";
            return Layer;
        }
    }
}
