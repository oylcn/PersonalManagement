using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using PersonalManagement.Web.Models;

namespace PersonalManagement.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
            trytrytryrty
            thghgh
            
        }
    }
}
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Linq.Mapping;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using Excel_Import_Cls.Entity;
using Microsoft.Office.Interop.Excel;
using SJF.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Globalization;
using System.Net.Mail;
using SJF.Business.BusinessRules;
using SJF.Business.Service;
using SJF.MsOffice.Excel;

namespace Excel_Import_Cls
{
    public class EI_Worker
    {
        private Microsoft.Office.Interop.Excel.Application EI_App;
        private Microsoft.Office.Interop.Excel.Workbook wb;
        private Microsoft.Office.Interop.Excel.Worksheet sheet;
        //private string template_path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Excel_Files\\";
        private string template_path = ConfigurationSettings.AppSettings["app_path"];
        private const int simulate_count = 5;   //kaç tane örnek olarak field ismi yazacağı.
        private int dikey_field_cnt, yatay_field_cnt, sabit_field_cnt, data_field_cnt;
        public EI_Worker()
        {

        }

        ~EI_Worker()
        {
            Close_Excel_Application(false);
        }

        private void Init_Excel_Application()
        {
            EI_App = new Microsoft.Office.Interop.Excel.Application();
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            EI_App.ScreenUpdating = false;  //dont forget to turn to false before saving
            EI_App.DisplayAlerts = false;  //dont forget to turn to false before saving
            EI_App.Visible = false;
            EI_App.ErrorCheckingOptions.NumberAsText = false;
            EI_App.ErrorCheckingOptions.BackgroundChecking = false;
        }

        private void Save_Excel_Application(string filename)
        {
            try
            {
                System.IO.File.Delete(filename);
            }
            catch (Exception ex2) { }
            Exception ex = null;
            EI_App.ScreenUpdating = true;
            try
            {
                //wb.SaveAs(filename, XlFileFormat.xlExcel9795, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlShared,
                //   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //wb.SaveCopyAs(filename);
            }
            catch (Exception _ex)
            {
                ex = _ex;
            }
            Close_Excel_Application(true);
            if (ex != null)
                throw ex;
        }

        private void Close_Excel_Application(bool _throw)
        {
            Exception ex = null;
            try
            {
                wb.Close(false, null, null);
                EI_App.Workbooks.Close();
                EI_App.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EI_App);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            }
            catch (Exception _ex)
            {
                ex = _ex;
            }
            sheet = null;
            wb = null;
            EI_App = null;
            GC.Collect();
            if (ex != null && _throw)
                throw ex;
        }

        public string Generate_Excel_Template(string username, DataBase dataBase)
        {
            int ret_val = 1;
            Exception ex = null;
            Init_Excel_Application();
            //try
            //{
            DataSet ds = CreateDataSet(dataBase);
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = dataBase.NameTable.Length >= 31 ? dataBase.NameTable.Substring(0, 30) : dataBase.NameTable;

            try
            {
                sheet.Cells.Select();
                sheet.Cells.NumberFormat = "@"; //Tarih için yeni eklendi.
            }
            catch (Exception ex1)
            { }

            //Sabitler
            int sabit_row = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='SABIT'"))
            {
                sheet.Cells[sabit_row, 1] = dr["FIELD_DESCRIPTION"];
                ((Excel.Range)sheet.Cells[sabit_row, 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                sabit_row++;
            }
            Excel.Range rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sabit_row - 1, 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            //Yataylar
            int yatay_row = 1;
            int yatay_max_col = 1;    //en sağ kolonun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'"))
            {
                int enlarged_by_below = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK>" + dr["RANK"]))
                    enlarged_by_below *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_upper = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK<" + dr["RANK"]))
                    enlarged_for_upper *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                for (int j = 0; j < enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt] = dr["FIELD_DESCRIPTION"];
                    sheet.get_Range(sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt],
                        sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1]).Merge(Type.Missing);
                    if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                }
                yatay_max_col = dikey_field_cnt + 1 + (enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1;
                yatay_row++;
            }
            rng = sheet.get_Range(sheet.Cells[1 + sabit_field_cnt, dikey_field_cnt + 1], sheet.Cells[yatay_row + sabit_field_cnt - 1, yatay_max_col]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            //Datalar
            int data_col = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'"))
            {
                sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1] = dr["FIELD_DESCRIPTION"];
                ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                data_col++;
            }

            //Dikeyler - biri diğerinin transpose'u
            int dikey_column = 1;
            int dikey_max_row = 1;    //en alt rowun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'"))
            {
                int enlarged_by_right = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK>" + dr["RANK"]))
                    enlarged_by_right *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_left = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK<" + dr["RANK"]))
                    enlarged_for_left *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                for (int j = 0; j < enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column] = dr["FIELD_DESCRIPTION"];
                    sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column],
                        sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right + enlarged_by_right - 1, dikey_column]).Merge(Type.Missing);
                    if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                }
                dikey_max_row = yatay_field_cnt + sabit_field_cnt + 1 + (enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_right + enlarged_by_right - 1;
                dikey_column++;
            }
            rng = sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, 1], sheet.Cells[dikey_max_row, dikey_column - 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.ColumnWidth = 5;
            //sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.NumberFormat = "@";	//!! bu varken sayıların ondalıklarını göstermiyordu - deneyelim bakalım
            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.AutoFit();

            /*}
            catch (Exception _ex)
            {
                ex = _ex;
            }*/
            string filename = @"..\Excel_Files\Schema\" + GetTableDefinitionByName(username, dataBase.NameTable).Replace('/', '_').Replace(' ', '_') + "_ADIM1_" + DateTime.Now.ToString("yyyyMMddhhmmss");
            Save_Excel_Application(template_path + filename + ".xlsx");
            //ParseExcel(table_name, filename);
            if (ex != null)
                throw ex;
            return filename;
        }


        public int ParseExcel(DataBase database, string filename, int boyut)
        {
            int ret_val = 1;
            int limit = 5000;
            switch (boyut)
            {
                case 5000:
                    limit = 5000; break;
                case 7500:
                    limit = 7500; break;
                case 20000:
                    limit = 20000; break;
                default:
                    limit = boyut > 0 ? boyut : limit;
                    break;
            }

            Exception ex = null;
            Init_Excel_Application();
            /*try
            {*/
            DataSet ds = CreateDataSet(database);
            wb = EI_App.Workbooks.Open(filename + ".xlsx", false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing);
            sheet = (Worksheet)wb.Sheets[1];
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");

            string cell_val = "", tmp = "";
            //Yatay
            int merge_col_cnt = 1;
            int max_col = 1;
            bool merged;
            for (int j = 0; j <= yatay_field_cnt; j++)	//= yoktu !!
            {
                for (int k = 0; k < 2000; k++)  //sağda en fazla bu kadar gider diyorum
                {
                    merged = false;
                    merge_col_cnt = 1;
                    if ((bool)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).MergeCells) //bu bir merged cell mi?
                    {
                        merge_col_cnt = ((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).MergeArea.Count; //merge sağda kaç kolon ilerliyor - bu kolonları unmerge edip, merge'in baş kolonundaki değerle dolduracağız.
                        cell_val = (string)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).Text;   //merge başı değeri
                        ((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).UnMerge();
                        for (int l = 0; l < merge_col_cnt - 1; l++) //merge'lü bölgenin merge'ünü kaldırınca, boş cell'leri, merge başı değeri ile doldur
                            sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt + l + 1] = cell_val;
                        merged = true;
                    }
                    k += merge_col_cnt - 1; //mergelü bölge kadar ilerle
                    tmp = (string)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).Text;
                    if (tmp != "")  //içinde değer olan max kolon..
                        if (k + 1 + dikey_field_cnt > max_col) max_col = k + 1 + dikey_field_cnt;
                        else if (!merged)
                            break;
                }
            }
            //Dikey
            int merge_row_cnt = 1;
            int max_row = 1;
            for (int j = 0; j < dikey_field_cnt; j++)
            {
                for (int k = 0; k < limit; k++)  //aşağıya en fazla bu kadar gider diyorum
                {
                    merged = false;
                    merge_row_cnt = 1;
                    if ((bool)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).MergeCells) //bu bir merged cell mi?
                    {
                        merge_row_cnt = ((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).MergeArea.Count; //merge sağda kaç kolon ilerliyor - bu kolonları unmerge edip, merge'in baş kolonundaki değerle dolduracağız.
                        cell_val = (string)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).Text;   //merge başı değeri
                        ((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).UnMerge();
                        for (int l = 0; l < merge_row_cnt - 1; l++) //merge'lü bölgenin merge'ünü kaldırınca, boş cell'leri, merge başı değeri ile doldur
                            sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt + l + 1, j + 1] = cell_val;
                        merged = true;
                    }
                    k += merge_row_cnt - 1; //mergelü bölge kadar ilerle
                    tmp = (string)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).Text;
                    if (tmp != "")  //içinde değer olan max row..
                        if (k + 1 + sabit_field_cnt + yatay_field_cnt > max_row) max_row = k + 1 + sabit_field_cnt + yatay_field_cnt;
                        else if (!merged)
                            break;
                }
            }

            //Başka bir sheet'e düz olarak tablo değerleri yazılıyor..
            Excel.Worksheet sheet2 = (Excel.Worksheet)wb.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            sheet2.Name = "NİHAİ DATA";
            try
            {
                sheet.Cells.Select();
                sheet.Cells.NumberFormat = "@"; //Tarih için yeni eklendi.
            }
            catch (Exception ex2)
            { }

            //sheet2.get_Range(sheet2.Cells[1, 1], sheet2.Cells[1, yatay_field_cnt + dikey_field_cnt + data_field_cnt + sabit_field_cnt + 2]).EntireColumn.NumberFormat = "@";	//!! bu varken sayıların ondalıklarını göstermiyordu - deneyelim bakalım
            int cur_col = 1, cur_row = 2;
            object[,] rawData = new object[max_row * max_col / data_field_cnt + 1, yatay_field_cnt + dikey_field_cnt + data_field_cnt + sabit_field_cnt + 2];   //bu size fazla fazla. ilk indexte, daha küçük çarpımlar olabilir (yatay, dikey vs)
            object[,] sourceData = new object[max_row + 1, max_col + 1];
            RangeToArray(sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sourceData.GetLength(0), sourceData.GetLength(1)]),
                ref sourceData);

            #region no_fill_if_null kaldırıldı
            /*Dictionary<string, string>[] fill_rule_dict = new Dictionary<string, string>[data_field_cnt + 1];   //her bir data_field için dictionary. elemanları: NO_FILL_IF_NULL 1 veya "", 1 ise; GECER_ORIENTATION - YATAY, DIKEY, SABIT. GECER_RANK - O orientation'da kaçıncı fielddan gecer_tarih değeri olarak faydalanılacağı
            int nf_cnt = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'", "RANK"))
            {   //boş cell'lerin bilgilerini max gecer_tarihtekilerden getirme opsiyonlarını oku
                fill_rule_dict[nf_cnt] = new Dictionary<string, string>();
                fill_rule_dict[nf_cnt].Add("NO_FILL_IF_NULL", Convert.ToString(dr["NO_FILL_IF_NULL"]));
                if (Convert.ToString(dr["NO_FILL_IF_NULL"]) != "")
                {
                    DataRow[] gecer_dr = ds.Tables["SCHEMA_DS"].Select("FIELD_NAME='" + Convert.ToString(dr["NO_FILL_IF_NULL_GECER_TARIH_FIELD_NAME"]) + "'");
                    fill_rule_dict[nf_cnt].Add("GECER_ORIENTATION", Convert.ToString(gecer_dr[0]["EXCEL_ORIENTATION"]));
                    fill_rule_dict[nf_cnt].Add("GECER_RANK", Convert.ToString(gecer_dr[0]["RANK"]));
                }
                nf_cnt++;
            }*/
            #endregion

            for (int r = sabit_field_cnt + yatay_field_cnt + 1; r <= max_row; r++)    //rowları ilerle
            {
                for (int c = 1; c <= max_col - data_field_cnt; c = c + data_field_cnt)  //kolonları ilerle
                {
                    try
                    {
                        for (int y = 1; y <= yatay_field_cnt; y++)  //yataydaki değerleri teker teker
                        {
                            rawData[cur_row, cur_col++] = sourceData[sabit_field_cnt + y, dikey_field_cnt + c];
                        }
                        for (int dik = 1; dik <= dikey_field_cnt; dik++)  //dikeydeki değerleri teker teker
                        {
                            rawData[cur_row, cur_col++] = sourceData[r, dik];
                        }
                        for (int s = 1; s <= sabit_field_cnt; s++)  //sabit fieldları al
                        {
                            rawData[cur_row, cur_col++] = sourceData[s, 1];
                        }
                        for (int dat = 1; dat <= data_field_cnt; dat++) //datanın bulunduğu yere git
                        {
                            rawData[cur_row, cur_col] = sourceData[r, dat + dikey_field_cnt + c - 1];
                            cur_col++;
                        }
                    }
                    catch (Exception exx)
                    {
                        cur_row--;
                    }
                    cur_row++;
                    cur_col = 1;
                }
            }

            sourceData = null; cur_col = 1;
            string tmp_data = "";
            //field descleri ile rowa koy ve kontrolleri yap.
            int comp_col = 1;
            int raw_max_col = rawData.GetLength(1) - 1;
            int control_start_row;
            control_start_row = 2;
            int rw = 0;
            bool second_pass = false;
            ArrayList interval_al = new ArrayList();
            do
            {
                interval_al.Clear();
                comp_col = 1;
                foreach (string str in new string[] { "YATAY", "DIKEY", "SABIT", "DATA" })
                    foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='" + str + "'", "RANK"))
                    {
                        rawData[1, comp_col] = dr["FIELD_DESCRIPTION"].ToString();
                        for (rw = control_start_row; rw < cur_row; rw++)   //data yazılan son rowa kadar kontrol et
                        {
                            tmp_data = Convert.ToString(rawData[rw, comp_col]);
                            if (dr["INTERVAL_CAN_EXIST"].ToString() != "" && Convert.ToInt32(dr["INTERVAL_CAN_EXIST"].ToString()) == 1 && tmp_data.IndexOf('-') > 0) // 20-30 şeklindeki değerler - şu an için sadece sayılar
                            {
                                string[] tmp_arr = tmp_data.Split('-');
                                rawData[rw, comp_col] = tmp_arr[0];  //orjinal yerine başlangıç değerini koy
                                if (CheckRulesforArray(ref rawData, rw, comp_col, raw_max_col, dr, second_pass))  //açılmış hali (başlangıç değeri ile) ruleları geçiyorsa çoğalt
                                {
                                    for (float i = Convert.ToSingle(tmp_arr[0]) + 1; i <= Convert.ToSingle(tmp_arr[1]); i++)    //çoğaltma işlemi - bu arraylist en sonda, rawdata'ya eklenecek
                                    {
                                        ArrayList al_row = new ArrayList();
                                        for (int c = 1; c < raw_max_col; c++)
                                            if (c != comp_col)
                                                al_row.Add(rawData[rw, c]);
                                            else
                                                al_row.Add(i);  //intervaldaki, bu tekrar için değerini koy, orjinal değer yerine
                                        interval_al.Add(al_row);    //her bir interval rowunu array liste ekle
                                    }
                                }
                            }
                            CheckRulesforArray(ref rawData, rw, comp_col, raw_max_col, dr, second_pass);
                        }
                        comp_col++;
                    }
                control_start_row = rw;
                cur_row = AddListToArray(ref rawData, ref interval_al, control_start_row - 1);  //additional interval rowlarını sonuna ekle
            } while (interval_al.Count > 0 && (second_pass = true));    //intervallardan gelen elemanlar olduğu sürece, ve second_pass'i true olarak set et-checkrulesforarray'de kullanılsın
            sheet2.get_Range(sheet2.Cells[1, 1], sheet2.Cells[rawData.GetLength(0), rawData.GetLength(1)]).NumberFormat = "@";

            ArrayToRange(ref rawData, sheet2.get_Range(sheet2.Cells[1, 1], sheet2.Cells[rawData.GetLength(0), rawData.GetLength(1)]));
            /*}
            catch (Exception _ex)
            {
                ex = _ex;
            }*/
            Save_Excel_Application(filename.Replace("_ADIM1", "") + "_ADIM2" + ".xlsx");
            //Close_Excel_Application(true);    şimdilik commentli !!
            if (ex != null)
                throw ex;
            return ret_val;
        }

        public string PutNihaDataToDBLive(DataBase database, string filename, string username, int taskID, int fileOrderNo, bool gecmisTarihIzinliMi, out int hataliKayitSayisi)
        {
            hataliKayitSayisi = 0;
            bool gecmisTarihIzinli = gecmisTarihIzinliMi;
            if (!gecmisTarihIzinli)
            {
                gecmisTarihIzinli = GecmisTarihliYuklemeyeIzniVarmi(database.NameTable);
            }
            Init_Excel_Application();
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            //System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = "CANLI ORTAMA AKTARILAN";

            try
            {
                sheet.Cells.Select();
                sheet.Cells.NumberFormat = "@"; //Tarih için yeni eklendi.
            }
            catch (Exception)
            {

            }


            System.Data.DataSet dsSchema = CreateDataSet(database);
            System.Data.DataTable dtSchema = dsSchema.Tables[0];
            System.Data.DataTable dtTestTable = GetDataFromTableName(database, taskID, fileOrderNo);
            int toplamKayitSayisi = dtTestTable.Rows.Count;


            var kolonAciklama = dtSchema.AsEnumerable().Select(dc => dc.Field<string>("FIELD_DESCRIPTION")).ToArray();
            var kolonArray = dtSchema.AsEnumerable().Select(dc => dc.Field<string>("FIELD_NAME"));
            var kolonIsimleri = string.Join(",", kolonArray.ToArray());

            int max_col = dikey_field_cnt + yatay_field_cnt + sabit_field_cnt + data_field_cnt + 1;
            object[,] rawData = new object[toplamKayitSayisi + 2, max_col + 1];
            for (int i = 0; i < kolonAciklama.Count(); i++)
            {
                rawData[1, i + 1] = kolonAciklama[i];
            }
            //ArrayToRange(ref rawData, sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[rawData.GetLength(0), rawData.GetLength(1)]));


            SqlConnection connection = new SqlConnection(database.ConnectionString);
            connection.Open();
            SqlTransaction transaction;
            transaction = connection.BeginTransaction("InsertTransaction");

            for (int i = 0; i < toplamKayitSayisi; i++)
            {
                string sql = "";
                string sql_params = "";
                bool gecerTarihEski = false;
                string error_message = "";
                int rows_inserted = 0;

                string sql_insert_prefix = "INSERT INTO " + string.Format("{0}.{1}.{2}", database.NameDatabase, database.NameSchema, database.NameTable) +
                                           " (LAST_USER_NAME, ENTRY_DATE, ENTRY_TIME,";
                sql_insert_prefix += kolonIsimleri + ") ";

                int columnIndex = 0;
                foreach (DataRow row in dtSchema.Rows)
                {
                    rawData[i + 2, columnIndex + 1] = dtTestTable.Rows[i][row["FIELD_NAME"].ToString()];
                    if (row["FIELD_TYPE"].ToString().ToLower().Contains("numer"))
                    {
                        sql_params += (dtTestTable.Rows[i][row["FIELD_NAME"].ToString()].ToString() == "" ? "NULL" : dtTestTable.Rows[i][row["FIELD_NAME"].ToString()].ToString().Replace(",", ".")) + ",";
                    }
                    else if (row["FIELD_TYPE"].ToString().ToLower().Contains("date"))
                    {
                        rawData[i + 2, columnIndex + 1] = Convert.ToDateTime(dtTestTable.Rows[i][row["FIELD_NAME"].ToString()].ToString()).ToString("dd.MM.yyyy");
                        sql_params += "CONVERT(DATETIME, '" + Convert.ToDateTime(dtTestTable.Rows[i][row["FIELD_NAME"].ToString()].ToString()).ToString("dd.MM.yyyy") + "', 104), ";
                    }
                    else
                    {
                        object deger = dtTestTable.Rows[i][row["FIELD_NAME"].ToString()] == null ? dtTestTable.Rows[i][row["FIELD_NAME"].ToString()] : dtTestTable.Rows[i][row["FIELD_NAME"].ToString()].ToString().Replace("\'", "''");
                        sql_params += "'" + deger + "',";
                    }

                    if (row["FIELD_NAME"].ToString() == "GECER_TARIH" && Convert.ToDateTime(dtTestTable.Rows[i][row["FIELD_NAME"].ToString()]) < DateTime.Today)
                    {
                        gecerTarihEski = true;
                    }
                    columnIndex++;
                }
                sql_params = sql_params.TrimEnd(',', ' ');

                sql_params =
                    string.Format(" VALUES ('{0}',CONVERT(DATETIME,CONVERT(VARCHAR, GETDATE(),104),104),'{1}',{2})",
                        username, DateTime.Now.ToString("HH:mm"), sql_params);

                sql = sql_insert_prefix + sql_params;


                SqlCommand command = new SqlCommand(sql, connection, transaction);


                try
                {
                    if (!gecmisTarihIzinli && gecerTarihEski)
                    {
                        throw new Exception("Geçer Tarih bugünden küçük kayıtlar canlı ortama aktarılamaz.");
                    }
                    rows_inserted = command.ExecuteNonQuery();
                }
                catch (Exception exc)
                {
                    error_message = exc.Message;
                    rows_inserted = 0;
                }

                if (rows_inserted != 1)
                {
                    rawData[i + 2, max_col] = " **HATA**INSERT**" + error_message;
                    error_message = "";
                    hataliKayitSayisi++;
                }
                if (i % 500 == 0)
                {
                    try
                    {
                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                    }
                    transaction = connection.BeginTransaction("InsertTransaction");
                    command.Transaction = transaction;
                }
            }
            try
            {
                transaction.Commit();
            }
            catch (Exception)
            {
            }
            finally
            {
                connection.Close();
            }
            ArrayToRange(ref rawData,
                    sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[rawData.GetLength(0) + 1, rawData.GetLength(1)]));

            Save_Excel_Application(filename);
            return filename;
        }

        private static int GetLastRow(Worksheet worksheet)
        {
            int lastUsedRow = 1;
            Range range = worksheet.UsedRange;
            for (int i = 1; i < range.Columns.Count; i++)
            {
                int lastRow = range.Rows.Count;
                for (int j = range.Rows.Count; j > 0; j--)
                {
                    if (lastUsedRow < lastRow)
                    {
                        lastRow = j;
                        if (!String.IsNullOrWhiteSpace(Convert.ToString((worksheet.Cells[j, i] as Range).Value)))
                        {
                            if (lastUsedRow < lastRow)
                                lastUsedRow = lastRow;
                            if (lastUsedRow == range.Rows.Count)
                                return lastUsedRow - 1;
                            break;
                        }
                    }
                    else
                        break;
                }
            }
            return lastUsedRow;
        }

        public void PutNihaiDataToDB(DataBase database, string filename, string username, int taskID, int ileriZamanGunSayisi, out int fileOrderNo, out DateTime gecerTarih, out int excelRowCount, out int hataliKayitSayisi)
        {
            Init_Excel_Application();
            try
            {
                hataliKayitSayisi = 0;
                DataSet ds = CreateDataSet(database);
                wb = EI_App.Workbooks.Open(filename + ".xlsx", false, false, Type.Missing, Type.Missing, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing,
                    Type.Missing);
                sheet = (Worksheet)wb.Sheets[1]; //nihai datanın bulunduğu sheet - ilk sheet
                //System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");


                excelRowCount = GetLastRow(sheet);

                Worksheet sheet2 = (Worksheet)wb.Sheets[2];
                DateTime tempGecerTarih = ConvertToDateTime(sheet2.get_Range("A1").Value2.ToString());
                gecerTarih = DateTime.ParseExact(tempGecerTarih.ToString("dd.MM.yyyy"), "dd.MM.yyyy", null);
                TableManager tableManager = new TableManager(taskID, database, gecerTarih);
                if (!tableManager.IsTableToDB())
                {
                    tableManager.CreateTable();
                }
                else
                {
                    tableManager.GecerTarihliKayitYoksaEkle(gecerTarih.ToString("yyyyMMdd"));
                }

                fileOrderNo = tableManager.GetFileOrderNo(taskID);

                int max_row = ((Excel.Range)sheet.Cells[1, 1]).get_End(XlDirection.xlDown).Row;
                int max_col = dikey_field_cnt + yatay_field_cnt + sabit_field_cnt + data_field_cnt + 1;
                //bir de hata kolonu
                object[,] rawData = new object[max_row, max_col + 1];
                RangeToArray(
                    sheet.get_Range(sheet.Cells[2, 1], sheet.Cells[rawData.GetLength(0) + 1, rawData.GetLength(1)]),
                    ref rawData);
                string sql_params = "", sql = "";

                string sql_insert_prefix = String.Format("INSERT INTO [{0}].[{1}].[{2}] (LAST_USER_NAME, ENTRY_DATE, ENTRY_TIME, TASK_ID, FILE_ORDER_NO,", database.NameDatabase, database.TempNameSchema, database.NameTable);
                int param = 1;
                List<int> data_params = new List<int>();
                foreach (string str in new string[] { "YATAY", "DIKEY", "SABIT", "DATA" })
                    foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='" + str + "'", "RANK"))
                    {
                        sql_insert_prefix += "[" + dr["FIELD_NAME"] + "], ";
                        if (!dr["FIELD_TYPE"].ToString().ToLower().Contains("datetime"))
                            sql_params += "@" + param.ToString() + ", ";
                        else
                            sql_params += "CONVERT(DATETIME, @" + param.ToString() + ", 104), ";
                        if (str == "DATA") data_params.Add(param);
                        param++;
                    }
                sql_insert_prefix = sql_insert_prefix.TrimEnd(new char[] { ',', ' ' });
                sql_insert_prefix += ") ";
                sql_params = sql_params.TrimEnd(new char[] { ',', ' ' });
                sql_params = "VALUES ('" + username + "', CONVERT(DATETIME,CONVERT(VARCHAR, GETDATE(),104),104), '" +
                             DateTime.Now.ToString("HH:mm") + "'," + taskID + "," + fileOrderNo + "," + sql_params + ")";
                sql = sql_insert_prefix + sql_params;
                SqlConnection connection = new SqlConnection(database.ConnectionString);
                connection.Open();
                SqlTransaction transaction;
                transaction = connection.BeginTransaction("MyTransaction");
                SqlCommand command = new SqlCommand(sql, connection, transaction);
                int rows_inserted = 0;
                string error_message = "";
                bool date_check = true;

                for (int r = 1; r < max_row; r++)
                {
                    if (Convert.ToString(rawData[r, max_col]).StartsWith(" **HATA**"))
                        continue;
                    param = 1;
                    date_check = true;
                    command.Parameters.Clear();
                    foreach (string str in new string[] { "YATAY", "DIKEY", "SABIT", "DATA" })
                    {
                        CultureInfo ci = new CultureInfo("tr-TR");

                        if (str == "SABIT")
                        {
                            DateTime dt = Convert.ToDateTime(rawData[r, param].ToString(), ci);
                            if ((dt - DateTime.Today).Days > ileriZamanGunSayisi)
                            {
                                date_check = false;
                                break;
                            }
                        }
                        foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='" + str + "'", "RANK"))
                        {
                            if (dr["PREFIX_STRING"] != null && dr["PREFIX_STRING"].ToString() != "" && rawData[r, param] != null)
                                rawData[r, param] = dr["PREFIX_STRING"].ToString() + rawData[r, param].ToString();
                            if (dr["SUFFIX_STRING"] != null && dr["SUFFIX_STRING"].ToString() != "" && rawData[r, param] != null)
                                rawData[r, param] = rawData[r, param].ToString() + dr["SUFFIX_STRING"].ToString();

                            if (dr["PAD_DIRECTION"] != null && dr["PAD_DIRECTION"].ToString() == "LEFT")
                                command.Parameters.AddWithValue("@" + param,
                                    rawData[r, param++].ToString()
                                        .PadLeft(Convert.ToInt32(dr["PAD_COUNT"]),
                                            dr["PADDING_CHAR"].ToString().ToCharArray()[0]));
                            else if (dr["PAD_DIRECTION"] != null && dr["PAD_DIRECTION"].ToString() == "RIGHT")
                                command.Parameters.AddWithValue("@" + param,
                                    rawData[r, param++].ToString()
                                        .PadRight(Convert.ToInt32(dr["PAD_COUNT"]),
                                            dr["PADDING_CHAR"].ToString().ToCharArray()[0]));
                            else
                            {
                                if (dr["FIELD_TYPE"].ToString().ToLower().StartsWith("numeric") && rawData[r, param] == null)
                                {
                                    command.Parameters.AddWithValue("@" + param++, DBNull.Value);
                                }
                                else if (dr["FIELD_TYPE"].ToString().ToLower().StartsWith("numeric"))
                                {
                                    command.Parameters.AddWithValue("@" + param, Convert.ToString(rawData[r, param++]).Replace(',', '.'));
                                }
                                else if (rawData[r, param] == null)
                                {
                                    command.Parameters.AddWithValue("@" + param++, DBNull.Value);
                                }
                                else
                                {
                                    command.Parameters.AddWithValue("@" + param, Convert.ToString(rawData[r, param++])).SqlDbType = System.Data.SqlDbType.VarChar;

                                }
                            }
                        }
                    }
                    bool full = false;
                    rows_inserted = 0;
                    if (date_check)
                    {
                        foreach (var item in data_params)
                        {
                            //if (command.Parameters["@" + item].Value == DBNull.Value || Convert.ToString(command.Parameters["@" + item].Value) != "")
                            //    full = full || true;
                            full = string.IsNullOrEmpty(command.Parameters["@" + item].Value + "") == false;
                            if (full) break;
                        }
                    }
                    if (full && date_check)
                    {
                        try
                        {
                            string commandQuery = SqlCommandDumper.GetCommandText(command);
                            rows_inserted = command.ExecuteNonQuery();
                        }
                        catch (Exception exc)
                        {
                            error_message = exc.Message;
                            rows_inserted = 0;
                        }
                    }
                    else if (date_check == false)
                    {
                        error_message = string.Format(" {0} GÜNDEN BÜYÜK GİRİLEMEZ", ileriZamanGunSayisi);
                        rows_inserted = 0;
                    }
                    else
                    {
                        error_message = " BOŞ DEĞER";
                        rows_inserted = 0;
                    }
                    if (rows_inserted != 1)
                    {
                        rawData[r, max_col] += " **HATA**INSERT**" + error_message;
                        error_message = "";
                        hataliKayitSayisi++;
                    }
                    if (r % 500 == 0)
                    {
                        try
                        {
                            transaction.Commit();
                        }
                        catch (Exception)
                        {
                        }
                        transaction = connection.BeginTransaction("MyTransaction");
                        command.Transaction = transaction;
                    }
                }
                try
                {
                    transaction.Commit();
                }
                catch (Exception exTransCommit)
                {
                    string message = exTransCommit.Message;
                }
                finally
                {
                    connection.Close();
                }
                ArrayToRange(ref rawData,
                    sheet.get_Range(sheet.Cells[2, 1], sheet.Cells[rawData.GetLength(0) + 1, rawData.GetLength(1)]));


                Save_Excel_Application(filename.Replace("_ADIM2", "") + "_ADIM3" + ".xlsx");
                //Close_Excel_Application(true);
            }
            catch (Exception exx)
            {
                Close_Excel_Application(true);
                throw exx;
            }

        }

        public string Generate_Excel_IkitarihArasi_Data(string username, DataBase database, string eklenme_bas_tarihi, string eklenme_bit_tarihi, int boyut)
        {
            Exception ex = null;
            Init_Excel_Application();
            DataSet ds = CreateDataSet(database);
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = database.NameTable.Length >= 23 ? database.NameTable.Substring(0, 22) + "_EKLENME" : database.NameTable + "_EKLENME";

            //Tablo datasını al
            DataSet data_ds = CreateTableEklenmeArasiDataSet(database, eklenme_bas_tarihi, eklenme_bit_tarihi, ds);   //bu sayede dateformat bozulmuyor..
            DataRow[] data_dr = new DataRow[data_ds.Tables[0].Rows.Count];
            if (data_dr.Length == 0)
                return "Hata: Seçtiğiniz tarih aralığında veri girişi yapılmamış!";
            data_ds.Tables[0].Rows.CopyTo(data_dr, 0);
            DisplayDataSetHeader(ds, database.NameTable, data_ds.Tables[0], sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, data_dr[0].ItemArray.Length]));
            DataRowArrayToRange(ref data_dr, sheet.get_Range(sheet.Cells[2, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]));

            Excel.Range rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;

            string filename = @"..\Excel_Files\Data\" + GetTableDefinitionByName(username, database.NameTable).Replace('/', '_') + "_EKLENENE_GORE_LISTE_VERI_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Save_Excel_Application(template_path + filename);
            if (ex != null)
                throw ex;
            return filename;
        }

        public string Generate_GecerlilikteListe_Data(string username, DataBase database, string eklenme_bas_tarihi, string eklenme_bit_tarihi, int boyut)
        {
            Exception ex = null;
            Init_Excel_Application();
            DataSet ds = CreateDataSet(database);
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = database.NameTable.Length >= 19 ? database.NameTable.Substring(0, 18) + "_GECERLILIK" : database.NameTable + "_GECERLILIK";
            //Tablo datasını al
            DataSet data_ds = CreateTableGecerlilikArasiDataSet(database, eklenme_bas_tarihi, eklenme_bit_tarihi, ds);   //bu sayede dateformat bozulmuyor..
            DataRow[] data_dr = new DataRow[data_ds.Tables[0].Rows.Count];
            if (data_dr.Length == 0)
                return "Hata: Seçtiğiniz tarih aralığında veri girişi yapılmamış!";
            data_ds.Tables[0].Rows.CopyTo(data_dr, 0);
            DisplayDataSetHeader(ds, database.NameTable, data_ds.Tables[0], sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, data_dr[0].ItemArray.Length]));
            DataRowArrayToRange(ref data_dr, sheet.get_Range(sheet.Cells[2, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]));

            Excel.Range rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;

            string filename = @"..\Excel_Files\Data\" + GetTableDefinitionByName(username, database.NameTable).Replace('/', '_') + "_GECERLILIGE_GORE_LISTE_VERI_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Save_Excel_Application(template_path + filename);
            if (ex != null)
                throw ex;
            return filename;
        }

        public string Generate_BelirliGecerlilikteListe_Data(string username, DataBase database, string gecerlilik_tarihi, int boyut)
        {
            Exception ex = null;
            Init_Excel_Application();
            DataSet ds = CreateDataSet(database);
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = database.NameTable.Length >= 19 ? database.NameTable.Substring(0, 18) + "_GECERLILIK" : database.NameTable + "_GECERLILIK";
            //Tablo datasını al
            DataSet data_ds = CreateTableDataSet(database, gecerlilik_tarihi, ds);   //bu sayede dateformat bozulmuyor..
            DataRow[] data_dr = new DataRow[data_ds.Tables[0].Rows.Count];
            if (data_dr.Length == 0)
                return "Hata: Seçtiğiniz tarih aralığında veri girişi yapılmamış!";
            data_ds.Tables[0].Rows.CopyTo(data_dr, 0);
            DisplayDataSetHeader(ds, database.NameTable, data_ds.Tables[0], sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, data_dr[0].ItemArray.Length]));
            DataRowArrayToRange(ref data_dr, sheet.get_Range(sheet.Cells[2, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]));

            Excel.Range rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[data_dr.Length + 1, data_dr[0].ItemArray.Length]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;

            string filename = @"..\Excel_Files\Data\" + GetTableDefinitionByName(username, database.NameTable).Replace('/', '_') + "_GECERLILIGE_GORE_LISTE_VERI_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Save_Excel_Application(template_path + filename);
            if (ex != null)
                throw ex;
            return filename;
        }

        public string Generate_Excel_NDim_Data(string username, DataBase database, string gecerlilik_tarihi, int boyut, string path = @"..\Excel_Files\Data\")
        {
            int ret_val = 1;
            int limit = 5000;
            switch (boyut)
            {
                case 5000:
                    limit = 5000; break;
                case 7500:
                    limit = 7500; break;
                case 20000:
                    limit = 20000; break;
                default:
                    limit = boyut > 0 ? boyut : limit;
                    break;
            }

            Init_Excel_Application();
            /*try
            {*/
            DataSet ds = CreateDataSet(database);
            wb = EI_App.Workbooks.Add(true);
            sheet = (Worksheet)wb.Sheets[1];
            //System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("tr-TR");
            sheet.Name = database.NameTable.Length >= 31 ? database.NameTable.Substring(0, 30) : database.NameTable;

            //Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType. xlCellTypeLastCell, Type.Missing); BURADA GEREKMİYOR AMA BELKİ EN SAĞDAKİ HÜCREYİ BULMAK İÇİN İŞE YARAYABİLİR (2. excel SATIRDA)
            try
            {
                sheet.Cells.Select();
                sheet.Cells.NumberFormat = "@"; //Tarih için yeni eklendi.
            }
            catch (Exception ex3)
            { }

            //Tablo datasını al
            DataSet data_ds = CreateTableDataSet(database, gecerlilik_tarihi, ds);


            //Sabitler
            int sabit_row = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='SABIT'"))
            {
                sheet.Cells[sabit_row, 1] = gecerlilik_tarihi;
                ((Excel.Range)sheet.Cells[sabit_row, 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"] + "\r\nNOT: Aşağıdaki dataların hepsinde aynı geçerlilik tarihi olmak zorunda değildir!");
                ((Excel.Range)sheet.Cells[sabit_row, 1]).Comment.Shape.Height = 150;
                sabit_row++;
            }
            Excel.Range rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sabit_row - 1, 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            if (yatay_field_cnt == 0)
            {
                return NDimDataDikey(username, database, ds, data_ds, rng, limit, sabit_row, gecerlilik_tarihi, path);
            }
            else
            {
                return NDimDataMatris(username, database, ds, data_ds, rng, limit, sabit_row, gecerlilik_tarihi, path);
            }

        }

        private string NDimDataMatris(string username, DataBase database, DataSet ds, DataSet data_ds, Excel.Range rng, int limit, int sabit_row, string gecerlilik_tarihi, string path)
        {
            Exception ex = null;
            //Yataylar
            int yatay_row = 1;
            int yatay_max_col = 1;    //en sağ kolonun yerini bil ki border ve alignment için
            bool yatayTanimVarMi = false;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'"))
            {
                int enlarged_by_below = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK>" + dr["RANK"]))
                    enlarged_by_below *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_upper = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK<" + dr["RANK"]))
                    enlarged_for_upper *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                DataRow[] distinct_data_dr = data_ds.Tables["DISTINCT_" + dr["FIELD_NAME"].ToString()].Select("", dr["FIELD_NAME"].ToString());	//bu fieldın distinct datalarını sorted olarak arraye koy ki headerlar bu arrayden yazabilsin
                for (int j = 0; j < enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    try
                    {
                        sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt] = distinct_data_dr[j % Convert.ToInt32(dr["VALUE_VARIATION_COUNT"])][0];
                    }
                    catch (Exception exx) { }
                    sheet.get_Range(sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt],
                        sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1]).Merge(Type.Missing);
                    if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                }
                yatay_max_col = dikey_field_cnt + 1 + (enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1;
                yatay_row++;
            }
            rng = sheet.get_Range(sheet.Cells[1 + sabit_field_cnt, dikey_field_cnt + 1], sheet.Cells[yatay_row + sabit_field_cnt - 1, yatay_max_col]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            //Datalar
            /*int data_col = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'"))
            {
                sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1] = dr["FIELD_DESCRIPTION"];
                ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                data_col++;
            }*/

            //Dikeyler - biri diğerinin transpose'u
            int dikey_column = 1;
            int dikey_max_row = 1;    //en alt rowun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'"))
            {
                int enlarged_by_right = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK>" + dr["RANK"]))
                    enlarged_by_right *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_left = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK<" + dr["RANK"]))
                    enlarged_for_left *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                DataRow[] distinct_data_dr = data_ds.Tables["DISTINCT_" + dr["FIELD_NAME"].ToString()].Select("", dr["FIELD_NAME"].ToString());	//bu fieldın distinct datalarını sorted olarak arraye koy ki headerlar bu arrayden yazabilsin
                for (int j = 0; j < enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    try
                    {
                        sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column] = distinct_data_dr[j % Convert.ToInt32(dr["VALUE_VARIATION_COUNT"])][0];
                    }
                    catch (Exception exx) { }
                    //sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column] = dr["FIELD_DESCRIPTION"];
                    sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column],
                        sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right + enlarged_by_right - 1, dikey_column]).Merge(Type.Missing);
                    if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                }
                dikey_max_row = yatay_field_cnt + sabit_field_cnt + 1 + (enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_right + enlarged_by_right - 1;
                dikey_column++;
            }
            //Buranın çıktısını test sonrası alacağız.
            //if (!yatayTanimVarMi)
            //{
            //    dikey_max_row = limit;
            //}

            rng = sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, 1], sheet.Cells[dikey_max_row, dikey_column - 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.ColumnWidth = 5;
            //sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.NumberFormat = "@";	//!! bu varken sayıların ondalıklarını göstermiyordu - deneyelim bakalım
            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.AutoFit();

            //
            //ParseExcel'den geliyor - Unmerge ve dataları oluşturma kısmı
            //
            string cell_val = "", tmp = "";
            //Yatay
            int merge_col_cnt = 1;
            int max_col = 1;
            bool merged;
            for (int j = 0; j <= yatay_field_cnt; j++)	//= yoktu !!
            {
                for (int k = 0; k < 2000; k++)  //sağda en fazla bu kadar gider diyorum
                {
                    merged = false;
                    merge_col_cnt = 1;
                    if ((bool)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).MergeCells) //bu bir merged cell mi?
                    {
                        merge_col_cnt = ((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).MergeArea.Count; //merge sağda kaç kolon ilerliyor - bu kolonları unmerge edip, merge'in baş kolonundaki değerle dolduracağız.
                        cell_val = (string)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).Text;   //merge başı değeri
                        ((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).UnMerge();
                        for (int l = 0; l < merge_col_cnt - 1; l++) //merge'lü bölgenin merge'ünü kaldırınca, boş cell'leri, merge başı değeri ile doldur
                            sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt + l + 1] = cell_val;
                        merged = true;
                    }
                    k += merge_col_cnt - 1; //mergelü bölge kadar ilerle
                    tmp = (string)((Excel.Range)sheet.Cells[j + 1 + sabit_field_cnt, k + 1 + dikey_field_cnt]).Text;
                    if (tmp != "")  //içinde değer olan max kolon..
                        if (k + 1 + dikey_field_cnt > max_col) max_col = k + 1 + dikey_field_cnt;
                        else if (!merged)
                            break;
                }
            }
            //Dikey
            int merge_row_cnt = 1;
            int max_row = 1;
            for (int j = 0; j < dikey_field_cnt; j++)
            {
                for (int k = 0; k < limit; k++)  //aşağıya en fazla bu kadar gider diyorum
                {
                    merged = false;
                    merge_row_cnt = 1;
                    if ((bool)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).MergeCells) //bu bir merged cell mi?
                    {
                        merge_row_cnt = ((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).MergeArea.Count; //merge sağda kaç kolon ilerliyor - bu kolonları unmerge edip, merge'in baş kolonundaki değerle dolduracağız.
                        cell_val = (string)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).Text;   //merge başı değeri
                        ((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).UnMerge();
                        for (int l = 0; l < merge_row_cnt - 1; l++) //merge'lü bölgenin merge'ünü kaldırınca, boş cell'leri, merge başı değeri ile doldur
                            sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt + l + 1, j + 1] = cell_val;
                        merged = true;
                    }
                    k += merge_row_cnt - 1; //mergelü bölge kadar ilerle
                    tmp = (string)((Excel.Range)sheet.Cells[k + 1 + sabit_field_cnt + yatay_field_cnt, j + 1]).Text;
                    if (tmp != "")  //içinde değer olan max row..
                        if (k + 1 + sabit_field_cnt + yatay_field_cnt > max_row) max_row = k + 1 + sabit_field_cnt + yatay_field_cnt;
                        else if (!merged)
                            break;
                }
            }
            //Unmerge bitti - şimdi datasetten her key için dataları bul ve yerine yaz
            int cur_col = 1, cur_row = 2;
            if (yatay_field_cnt == 0)	//!! hack - yatay fieldlar olmayınca üstte bir bölgeye girmediği için yatay_max_col'u tekrar hesaplıyorum
                yatay_max_col = dikey_field_cnt + 1 + data_field_cnt - 1;
            max_col = yatay_max_col;	//bu aldığım yere göre farklı
            object[,] sourceData = new object[max_row + 1, max_col + 1];
            RangeToArray(sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sourceData.GetLength(0), sourceData.GetLength(1)]),
                ref sourceData);
            string filter = "";
            int field_count = 0;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'"))
                if (dr["FIELD_TYPE"].ToString().StartsWith("NUMERIC"))
                    filter += dr["FIELD_NAME"].ToString() + "=@" + (++field_count).ToString() + " AND ";
                else
                {
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += "TRIM(";
                    filter += dr["FIELD_NAME"].ToString();	//datatable'da ' 2' olarak olmasına rağmen rangetoarray rng.value2'yi kopyalarken sanki numericmiş gibi kopyalayıp sourcedata'ya bunu 2 olarak atıyor. o zaman da bulamıyorduk. bulabilmek için aradığımız k1'i trim(k1)='2' şeklinde arıyor olacağız
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += ")";
                    filter += "=";
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += "TRIM(";
                    filter += "'@" + (++field_count).ToString() + "'";
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += ")";
                    filter += " AND ";
                    //filter += "='@" + (++field_count).ToString() + "' AND ";
                }
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'"))
                if (dr["FIELD_TYPE"].ToString().StartsWith("NUMERIC"))
                    filter += dr["FIELD_NAME"].ToString() + "=@" + (++field_count).ToString() + " AND ";
                else
                {
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += "TRIM(";
                    filter += dr["FIELD_NAME"].ToString();
                    if (dr["PAD_DIRECTION"].ToString() != "")	//datatable'da ' 2' olarak olmasına rağmen rangetoarray rng.value2'yi kopyalarken sanki numericmiş gibi kopyalayıp sourcedata'ya bunu 2 olarak atıyor. o zaman da bulamıyorduk. bulabilmek için aradığımız k1'i trim(k1)='2' şeklinde arıyor olacağız
                        filter += ")";
                    filter += "=";
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += "TRIM(";
                    filter += "'@" + (++field_count).ToString() + "'";
                    if (dr["PAD_DIRECTION"].ToString() != "")
                        filter += ")";
                    filter += " AND ";
                }
            filter = filter.TrimEnd(" AND ".ToCharArray());
            string[] data_field_names = new string[data_field_cnt];
            int dat_field_ind = 0;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'"))
                data_field_names[dat_field_ind++] = dr["FIELD_NAME"].ToString();
            DataRow[] select_dr;
            object data_obj;
            for (int r = sabit_field_cnt + yatay_field_cnt + 1; r <= max_row; r++)    //rowları ilerle
            {
                for (int c = 1; c <= max_col - data_field_cnt; c = c + data_field_cnt)  //kolonları ilerle
                {
                    string my_filter = filter;
                    string tmp_d = "";
                    int s_r, s_c;
                    for (s_r = 1; s_r <= yatay_field_cnt; s_r++)
                    {
                        try
                        {
                            tmp_d = Convert.ToString(sourceData[sabit_field_cnt + s_r, dikey_field_cnt + c]);	//!! ondalıklı sayılar patlayabilir
                        }
                        catch (Exception) { }
                        if (tmp_d == "")
                            goto label_next;
                        my_filter = my_filter.Replace("@" + s_r.ToString(), tmp_d);
                    }
                    for (s_c = 1; s_c <= dikey_field_cnt; s_c++)
                    {
                        tmp_d = Convert.ToString(sourceData[sabit_field_cnt + r - 1, s_c]);
                        if (tmp_d == "" || tmp_d.Trim() == "")
                            goto label_next;
                        my_filter = my_filter.Replace("@" + (s_r - 1 + s_c).ToString(), tmp_d);
                    }
                    for (int dat = 1; dat <= data_field_cnt; dat++) //datanın bulunduğu yere git
                    {
                        select_dr = data_ds.Tables["TABLE_DATA_DS"].Select(my_filter);
                        if (select_dr.Length > 0 && (data_obj = select_dr[0][data_field_names[dat - 1]]) != null)
                        {
                            try
                            {
                                sourceData[r, dat + dikey_field_cnt + c - 1] = data_obj; 	//!! ondalıklı sayılar patlayabilir
                            }
                            catch { }
                        }
                        cur_col++;
                    }
                label_next:
                    cur_row++;
                    cur_col = 1;
                }
            }
            ArrayToRange(ref sourceData, sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sourceData.GetLength(0), sourceData.GetLength(1)]));

            #region Bir daha merged cell oluştur
            //Sabitler
            sabit_row = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='SABIT'"))
            {
                sheet.Cells[sabit_row, 1] = gecerlilik_tarihi;
                //((Excel.Range)sheet.Cells[sabit_row, 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"] + "\r\nNOT: Aşağıdaki dataların hepsinde aynı geçerlilik tarihi olmak zorunda değildir!");
                //((Excel.Range)sheet.Cells[sabit_row, 1]).Comment.Shape.Height = 150;
                sabit_row++;
            }
            rng = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sabit_row - 1, 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra) - YOKTU

            //Yataylar
            yatay_row = 1;
            yatay_max_col = 1;    //en sağ kolonun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'"))
            {
                int enlarged_by_below = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK>" + dr["RANK"]))
                    enlarged_by_below *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_upper = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY' AND RANK<" + dr["RANK"]))
                    enlarged_for_upper *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                DataRow[] distinct_data_dr = data_ds.Tables["DISTINCT_" + dr["FIELD_NAME"].ToString()].Select("", dr["FIELD_NAME"].ToString());	//bu fieldın distinct datalarını sorted olarak arraye koy ki headerlar bu arrayden yazabilsin
                for (int j = 0; j < enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    try
                    {
                        sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt] = distinct_data_dr[j % Convert.ToInt32(dr["VALUE_VARIATION_COUNT"])][0];
                    }
                    catch (Exception exx) { }
                    sheet.get_Range(sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt],
                        sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1]).Merge(Type.Missing);
                    /*if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_row + sabit_field_cnt, dikey_field_cnt + 1 + j * enlarged_by_below * data_field_cnt]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);*/
                }
                yatay_max_col = dikey_field_cnt + 1 + (enlarged_for_upper * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_below * data_field_cnt + enlarged_by_below * data_field_cnt - 1;
                yatay_row++;
            }
            rng = sheet.get_Range(sheet.Cells[1 + sabit_field_cnt, dikey_field_cnt + 1], sheet.Cells[yatay_row + sabit_field_cnt - 1, yatay_max_col]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra) - YOKTU

            //Datalar
            /*int data_col = 1;
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'"))
            {
                sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1] = dr["FIELD_DESCRIPTION"];
                ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, dikey_field_cnt + 1 + data_col - 1]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                data_col++;
            }*/

            //Dikeyler - biri diğerinin transpose'u
            dikey_column = 1;
            dikey_max_row = 1;    //en alt rowun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'"))
            {
                int enlarged_by_right = 1;
                foreach (DataRow dr_below in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK>" + dr["RANK"]))
                    enlarged_by_right *= Convert.ToInt32(dr_below["VALUE_VARIATION_COUNT"]);
                int enlarged_for_left = 1;
                foreach (DataRow dr_upper in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' AND RANK<" + dr["RANK"]))
                    enlarged_for_left *= Convert.ToInt32(dr_upper["VALUE_VARIATION_COUNT"]);
                DataRow[] distinct_data_dr = data_ds.Tables["DISTINCT_" + dr["FIELD_NAME"].ToString()].Select("", dr["FIELD_NAME"].ToString());	//bu fieldın distinct datalarını sorted olarak arraye koy ki headerlar bu arrayden yazabilsin
                for (int j = 0; j < enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]); j++)
                {
                    try
                    {
                        sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column] = distinct_data_dr[j % Convert.ToInt32(dr["VALUE_VARIATION_COUNT"])][0];
                    }
                    catch (Exception exx) { }
                    //sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column] = dr["FIELD_DESCRIPTION"];
                    sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column],
                        sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right + enlarged_by_right - 1, dikey_column]).Merge(Type.Missing);
                    /*if (j == 0)
                        ((Excel.Range)sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1 + j * enlarged_by_right, dikey_column]).AddComment(dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);*/
                }
                dikey_max_row = yatay_field_cnt + sabit_field_cnt + 1 + (enlarged_for_left * Convert.ToInt32(dr["VALUE_VARIATION_COUNT"]) - 1) * enlarged_by_right + enlarged_by_right - 1;
                dikey_column++;
            }
            rng = sheet.get_Range(sheet.Cells[yatay_field_cnt + sabit_field_cnt + 1, 1], sheet.Cells[dikey_max_row, dikey_column - 1]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra) - YOKTU

            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.ColumnWidth = 5;
            //sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.NumberFormat = "@";	//!! bu varken sayıların ondalıklarını göstermiyordu - deneyelim bakalım
            sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[1, yatay_max_col]).EntireColumn.AutoFit();
            #endregion
            /*	}
			catch (Exception _ex)
			{
				ex = _ex;
			}*/
            string filename = path + GetTableDefinitionByName(username, database.NameTable).Replace('/', '_') + "_MEVCUT_VERI_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Save_Excel_Application(template_path + filename);
            //ParseExcel(table_name, filename);
            if (ex != null)
                throw ex;
            return filename;
        }

        private string NDimDataDikey(string username, DataBase database, DataSet ds, DataSet data_ds, Excel.Range rng, int limit, int sabit_row, string gecerlilik_tarihi, string path)
        {
            Exception ex = null;
            //Dikeyler - biri diğerinin transpose'u
            int dikey_column = 1;
            int dikey_max_row = 1;    //en alt rowun yerini bil ki border ve alignment için
            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'"))
            {
                for (int j = sabit_field_cnt; j <= data_ds.Tables["TABLE_DATA_DS"].Rows.Count; j++)
                {
                    sheet.Cells[j + sabit_field_cnt, dr["RANK"]] = data_ds.Tables["TABLE_DATA_DS"].Rows[j - 1][dr["FIELD_NAME"].ToString()];
                    if (j == sabit_field_cnt)
                    {
                        ((Excel.Range)sheet.Cells[j + sabit_field_cnt, dr["RANK"]]).AddComment(dr["FIELD_NAME"].ToString() + " : " + dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                    }
                    dikey_column = Convert.ToInt32(dr["RANK"]);
                }
                dikey_max_row = data_ds.Tables["TABLE_DATA_DS"].Rows.Count;
            }

            foreach (DataRow dr in ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'"))
            {
                for (int j = sabit_field_cnt; j <= data_ds.Tables["TABLE_DATA_DS"].Rows.Count; j++)
                {
                    sheet.Cells[j + sabit_field_cnt, dikey_column + Convert.ToInt32(dr["RANK"])] = FormatText(data_ds.Tables["TABLE_DATA_DS"].Rows[j - 1][dr["FIELD_NAME"].ToString()]);
                    if (j == sabit_field_cnt)
                    {
                        ((Excel.Range)sheet.Cells[j + sabit_field_cnt, dikey_column + Convert.ToInt32(dr["RANK"])]).AddComment(dr["FIELD_NAME"].ToString() + " : " + dr["FIELD_DESCRIPTION"] + " : " + dr["FIELD_TYPE"]);
                    }
                }
            }
            rng = sheet.get_Range(sheet.Cells[sabit_field_cnt + 1, 1], sheet.Cells[sabit_field_cnt + dikey_max_row, dikey_column]);
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.VerticalAlignment = XlHAlign.xlHAlignCenter;
            rng.NumberFormat = "@"; //yeni geldi tarih formatında gözükmesin 44084 yazmasın diye (edit edildikten sonra)

            //int cur_col = 1, cur_row = 2;
            //object[,] sourceData = new object[dikey_max_row, dikey_column + data_field_cnt];
            //RangeToArray(sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sourceData.GetLength(0), sourceData.GetLength(1)]), ref sourceData);

            //ArrayToRange(ref sourceData, sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[sourceData.GetLength(0), sourceData.GetLength(1)]));

            string filename = path + GetTableDefinitionByName(username, database.NameTable).Replace('/', '_') + "_MEVCUT_VERI_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx";
            Save_Excel_Application(template_path + filename);
            //ParseExcel(table_name, filename);
            if (ex != null)
                throw ex;
            return filename;
        }

        private object FormatText(object value)
        {
            if (value == null)
            {
                return value;
            }
            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.Boolean:
                    return Convert.ToBoolean(value);
                case TypeCode.Byte:
                    return Convert.ToByte(value);
                case TypeCode.Char:
                    return Convert.ToChar(value);
                case TypeCode.DateTime:
                    return Convert.ToDateTime(value.ToString()).ToString("dd.MM.yyyy");
                case TypeCode.Decimal:
                    return Convert.ToDecimal(value);
                case TypeCode.Double:
                    return Convert.ToDouble(value);
                case TypeCode.Empty:
                    throw new NullReferenceException("The target type is null.");
                case TypeCode.Int16:
                    return Convert.ToInt16(value);
                case TypeCode.Int32:
                    return Convert.ToInt32(value);
                case TypeCode.Int64:
                    return Convert.ToInt64(value);
                case TypeCode.Object:
                    // Leave conversion of non-base types to derived classes.
                    return value;
                case TypeCode.SByte:
                    return Convert.ToSByte(value);
                case TypeCode.Single:
                    return Convert.ToSingle(value);
                case TypeCode.String:
                    return Convert.ToString(value);
                case TypeCode.UInt16:
                    return Convert.ToUInt16(value);
                case TypeCode.UInt32:
                    return Convert.ToUInt32(value);
                case TypeCode.UInt64:
                    return Convert.ToUInt64(value);
                case TypeCode.DBNull:
                    return value;
                default:
                    throw new InvalidCastException("Conversion not supported.");
            }

        }

        private void RangeToArray(Excel.Range rng, ref object[,] destarray)
        {   //Excel'den kopyaladığında array 0. indexten başladığı için excel[1,1] --> array[0,0]. Array, excel row ve kolonlarını birebir temsil etsin diye 1 aşağı-sağa shift ediyoruz.
            int dim1 = destarray.GetLength(0);
            int dim2 = destarray.GetLength(1);
            object[,] tmparray = new object[dim1, dim2];
            Array.Copy((Array)rng.Value2, tmparray, dim1 * dim2);
            for (int i = 0; i < dim1 - 1; i++)
                for (int j = 0; j < dim2 - 1; j++)
                    destarray[i + 1, j + 1] = tmparray[i, j];
        }

        private void ArrayToRange(ref object[,] arr, Excel.Range rng)
        {   //Excel'e kopyalanırken array 0. indexinden assignment başladığı için, excel[1,1] <-- array[0,0] şekline getirebilmek için sola-yukarı shift edilmiş array kullanılıyor
            int dim1 = arr.GetLength(0);
            int dim2 = arr.GetLength(1);
            object[,] tmparr = new object[dim1, dim2];
            for (int i = 0; i < dim1 - 1; i++)
                for (int j = 0; j < dim2 - 1; j++)
                    tmparr[i, j] = arr[i + 1, j + 1];
            rng.Value2 = tmparr;
        }

        private void DisplayDataSetHeader(DataSet table_def_ds, string table_name, System.Data.DataTable tbl, Excel.Range rng)
        {
            List<string> ignore_columns = new List<string>(new string[] { "RELEASE_NO", "LAST_USER_NAME", "ENTRY_DATE", "ENTRY_TIME", "OPER_SFS_PRODUCT_NO", "OPER_NO" });
            int i = 1;
            foreach (System.Data.DataColumn col in tbl.Columns)
            {
                rng[1, i] = col.ColumnName;
                if (!ignore_columns.Contains(col.ColumnName))
                    ((Excel.Range)rng[1, i]).AddComment(table_def_ds.Tables["SCHEMA_DS"].Select("TABLE_NAME='" + table_name + "' AND FIELD_NAME='" + col.ColumnName + "'")[0]["FIELD_DESCRIPTION"]);
                i++;
            }

        }

        private void DataRowArrayToRange(ref DataRow[] arr, Excel.Range rng)
        {   //Excel'e kopyalanırken array 0. indexinden assignment başladığı için, excel[1,1] <-- array[0,0] şekline getirebilmek için sola-yukarı shift edilmiş array kullanılıyor

            int dim1 = rng.Cells.Rows.Count;
            int dim2 = rng.Cells.Columns.Count;
            object[,] tmparr = new object[dim1, dim2];
            for (int i = 0; i < dim1; i++)
                for (int j = 0; j < dim2; j++)
                    //tmparr[i, j] = arr[i][j];
                    rng[i + 1, j + 1] = arr[i][j];
            //rng.Value2 = tmparr;
        }


        private bool CheckRulesforArray(ref object[,] rawData, int row, int column, int raw_max_col/*her çağırılışta hesaplanmasın diye alıyoruz-burada hata var-yok tutuluyor*/
            , DataRow rule_dr, bool second_pass)
        {
            bool ret_val = true;
            string tmp_data = Convert.ToString(rawData[row, column]);
            if (second_pass && tmp_data.StartsWith(" **HATA**"))   //bu cell interval ile çoğaltılmış bir satırdan geliyor olabilir ve Hata işareti cell'e konmuş olabilir, bu durumda false dön ve başka bir kontrol yapma (second pass'i de koydum ki, ilk pass'te boş yere starts_with bakmasın)
                ret_val = false;
            else
            {
                if (rule_dr["MIN_VALUE"].ToString() != "" && CompareLT(tmp_data, rule_dr["MIN_VALUE"].ToString()))
                {
                    rawData[row, column] = " **HATA**MIN_VALUE**" + tmp_data;
                    ret_val = false;
                }
                if (rule_dr["MAX_VALUE"].ToString() != "" && CompareGT(tmp_data, rule_dr["MAX_VALUE"].ToString()))
                {
                    rawData[row, column] = " **HATA**MAX_VALUE**" + tmp_data;
                    ret_val = false;
                }
                if (rule_dr["VALUE_SET"].ToString() != "" && !InValueSet(rule_dr["VALUE_SET"].ToString(), tmp_data))
                {
                    rawData[row, column] = " **HATA**VALUE_SET**" + tmp_data;
                    ret_val = false;
                }
            }
            if (ret_val == false)
                rawData[row, raw_max_col] = " **HATA**";
            return ret_val;
        }

        private int AddListToArray(ref object[,] rawData, ref ArrayList addition_list, int add_after_index)
        {   //add_ettikten sonra gelinen son row'u döner (control_start_row kullanılmak üzere)
            int row = add_after_index + 1;
            if (addition_list.Count <= 0)
                return row;
            object[,] tmparr = new object[rawData.GetLength(0) + addition_list.Count, rawData.GetLength(1)];    //addition_listtekileri alacak kadar büyüt
            Array.Copy(rawData, tmparr, rawData.GetLength(0) * rawData.GetLength(1));
            foreach (ArrayList al_row in addition_list)
            {
                int addition_col = 1;
                foreach (object al_row_col in al_row)
                {
                    tmparr[row, addition_col++] = al_row_col;   //addition listteki her bir rowun her kolonunu ekle
                }
                row++;
            }
            rawData = null;
            rawData = new object[tmparr.GetLength(0), tmparr.GetLength(1)];
            Array.Copy(tmparr, rawData, tmparr.GetLength(0) * tmparr.GetLength(1));
            return row;
        }

        private bool CompareLT(string v1, string v2)
        {
            float num1, num2;
            if (Single.TryParse(v1, out num1))
                if (Single.TryParse(v2, out num2))
                    if (num1 < num2)    //numeric durum
                        return true;
                    else
                        return false;
            if (String.Compare(v1, v2, false) < 0)  //string durum - case sensitive
                return true;
            else
                return false;
        }

        private bool CompareGT(string v1, string v2)
        {
            float num1, num2;
            if (Single.TryParse(v1, out num1))
                if (Single.TryParse(v2, out num2))
                    if (num1 > num2)    //numeric durum
                        return true;
                    else
                        return false;
            if (String.Compare(v1, v2, false) > 0)  //string durum - case sensitive
                return true;
            else
                return false;
        }

        private bool InValueSet(string value_set, string v)
        {
            if (Array.IndexOf(value_set.Split(','), v) > -1)
                return true;
            else
                return false;
        }

        private DataSet CreateDataSet(DataBase database)
        {
            string sql = "";
            SqlConnection connection = new SqlConnection(Connections.WinsureConnectionString);
            connection.Open();
            sql = "SELECT * FROM fiba.EI_SCHEMA WHERE TABLE_NAME='" + database.NameTable + "' AND DELETED=0 ORDER BY RANK";
            SqlDataAdapter da = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "SCHEMA_DS");
            connection.Close();

            DataRow[] dikey = ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'");
            dikey.OrderBy(x => x.Field<int>("RANK"));

            var sabit = ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'");
            sabit.OrderBy(x => x.Field<int>("RANK"));

            var yatay = ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='SABIT'");
            yatay.OrderBy(x => x.Field<int>("RANK"));

            var data = ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'");
            data.OrderBy(x => x.Field<int>("RANK"));

            var tumTablo = dikey.Union(sabit).Union(yatay).Union(data);
            System.Data.DataTable table = new System.Data.DataTable();
            table.TableName = "SCHEMA_DS";

            foreach (DataColumn dataColumn in ds.Tables["SCHEMA_DS"].Columns)
            {
                table.Columns.Add(dataColumn.ColumnName, dataColumn.DataType);
            }
            foreach (var dataRow in tumTablo)
            {
                table.ImportRow(dataRow);
            }
            DataSet ds2 = new DataSet();
            ds2.Tables.Add(table);

            dikey_field_cnt = ds2.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY'").Length;
            yatay_field_cnt = ds2.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='YATAY'").Length;
            sabit_field_cnt = ds2.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='SABIT'").Length;
            data_field_cnt = ds2.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DATA'").Length;
            return ds2;
        }

        private DataSet CreateTableEklenmeArasiDataSet(DataBase database, string eklenme_bas_tarihi, string eklenme_bit_tarihi, DataSet table_def_ds)
        {
            //tablonun belli bir geçerlilik tarihindeki tüm geçerli datası
            string sql = "";
            SqlConnection connection = new SqlConnection(database.ConnectionString);
            connection.Open();
            List<string> ignore_columns = new List<string>(new string[] { "RELEASE_NO", "LAST_USER_NAME", "ENTRY_DATE", "ENTRY_TIME", "OPER_SFS_PRODUCT_NO", "OPER_NO" });
            sql = "SELECT * FROM " + string.Format("{0}.{1}.{2}", database.NameDatabase, database.NameSchema, database.NameTable) + " (NOLOCK) T WHERE ENTRY_DATE >= CONVERT(DATETIME, '" + eklenme_bas_tarihi + "', 104) AND ENTRY_DATE <= CONVERT(DATETIME, '" + eklenme_bit_tarihi + "', 104) + 1"
                + " ORDER BY ENTRY_DATE";
            foreach (DataRow dr in table_def_ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' OR EXCEL_ORIENTATION='YATAY'"))
                if (!ignore_columns.Contains(dr["FIELD_NAME"].ToString()))
                    sql += ", " + dr["FIELD_NAME"].ToString();
            SqlDataAdapter da = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "TABLE_DATA_DS");
            connection.Close();
            return ds;
        }

        private DataSet CreateTableGecerlilikArasiDataSet(DataBase database, string eklenme_bas_tarihi, string eklenme_bit_tarihi, DataSet table_def_ds)
        {	//tablonun belli bir geçerlilik tarihindeki tüm geçerli datası
            string sql = "";
            SqlConnection connection = new SqlConnection(database.ConnectionString);
            connection.Open();
            List<string> ignore_columns = new List<string>(new string[] { "RELEASE_NO", "LAST_USER_NAME", "ENTRY_DATE", "ENTRY_TIME", "OPER_SFS_PRODUCT_NO", "OPER_NO" });
            sql = "SELECT * FROM " + string.Format("{0}.{1}.{2}", database.NameDatabase, database.NameSchema, database.NameTable) + " (NOLOCK) T WHERE GECER_TARIH >= CONVERT(DATETIME, '" + eklenme_bas_tarihi + "', 104) AND GECER_TARIH <= CONVERT(DATETIME, '" + eklenme_bit_tarihi + "', 104) + 1"
                + " ORDER BY GECER_TARIH";
            foreach (DataRow dr in table_def_ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' OR EXCEL_ORIENTATION='YATAY'"))
                if (!ignore_columns.Contains(dr["FIELD_NAME"].ToString()))
                    sql += ", " + dr["FIELD_NAME"].ToString();
            SqlDataAdapter da = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "TABLE_DATA_DS");
            connection.Close();
            return ds;
        }

        private DataSet CreateTableDataSet(DataBase database, string gecerlilik_tarihi, DataSet table_def_ds)
        {	//tablonun belli bir geçerlilik tarihindeki tüm geçerli datası
            string sql = "";
            SqlConnection connection = new SqlConnection(database.ConnectionString);
            connection.Open();
            /*sql = "SELECT TOP 1 * FROM " + table_name + " (NOLOCK)";
            SqlCommand command = new SqlCommand(sql, connection);
            SqlDataReader reader = command.ExecuteReader(CommandBehavior.SingleRow);
            if (reader.Read())
            {
                for (byte i = 0; i < reader.FieldCount; i++)
                    police_bilgiler.Add(reader.GetName(i), reader.GetValue(i));
            }*/
            List<string> ignore_columns = new List<string>(new string[] { "RELEASE_NO", "LAST_USER_NAME", "ENTRY_DATE", "ENTRY_TIME", "OPER_SFS_PRODUCT_NO", "OPER_NO" });
            sql = "SELECT * FROM " + string.Format("{0}.{1}", database.NameSchema, database.NameTable) + " (NOLOCK) T WHERE GECER_TARIH="
                + " (SELECT MAX(GECER_TARIH) FROM " + string.Format("{0}.{1}", database.NameSchema, database.NameTable) + " (NOLOCK) WHERE ";
            foreach (DataRow dr in table_def_ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' OR EXCEL_ORIENTATION='YATAY'"))
                if (!ignore_columns.Contains(dr["FIELD_NAME"].ToString()))
                    sql += dr["FIELD_NAME"].ToString() + " = T." + dr["FIELD_NAME"].ToString() + " AND ";
            sql = sql.TrimEnd(" AND ".ToCharArray());
            sql += " AND GECER_TARIH <= CONVERT(DATETIME, '" + gecerlilik_tarihi + "', 104))";
            SqlDataAdapter da = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "TABLE_DATA_DS");
            connection.Close();

            //Distinctler
            foreach (DataRow dr in table_def_ds.Tables["SCHEMA_DS"].Select("EXCEL_ORIENTATION='DIKEY' OR EXCEL_ORIENTATION='YATAY'"))
                if (!ignore_columns.Contains(dr["FIELD_NAME"].ToString()))
                {
                    System.Data.DataTable dt = SelectDistinct(ds.Tables["TABLE_DATA_DS"], new string[] { dr["FIELD_NAME"].ToString() });
                    dt.TableName = "DISTINCT_" + dr["FIELD_NAME"].ToString();
                    ds.Tables.Add(dt);
                }
            return ds;
        }

        public DataSet GetTableDefs(string username)
        {
            username = SecurityHelper.FilterSQLInjection(username);
            using (SqlConnection connection = new SqlConnection(Connections.WinsureConnectionStringReadApp))
            {
                connection.Open();
                string sql =
                   string.Format(@"SELECT DISTINCT D.TABLE_NAME,
                               	               D.TABLE_DESCRIPTION,
                               	               CATEGORY,
                               	               D.GROUP_ID,
                               	               D.SERVERNAME,
                               	               D.DATABASENAME,
                               	               D.SCHEMANAME
                               FROM EI_TABLE_DEF D
                               INNER JOIN [10.81.24.180].FINANS.tableAuth.TABLE_AUTH_DETAIL T(NOLOCK) ON T.TABLE_NAME = D.TABLE_NAME
                               INNER JOIN [10.81.24.180].FINANS.tableAuth.TABLE_AUTH_AUTHORIZE_FOR_EXCEL_TARIFE TA(NOLOCK) ON TA.TABLE_AUTH_DETAIL_ID = T.ID
                               WHERE D.DELETED = 0
                               	AND TA.AUTHORIZE_FOR_EXCEL_TARIFE = 1
                               	AND TA.SICIL_NO = '{0}'", username);

                SqlDataAdapter da = new SqlDataAdapter(sql, connection);
                DataSet ds = new DataSet();
                da.Fill(ds, "EI_TABLE_DEF");
                connection.Close();

                //Distinctler
                foreach (DataColumn dc in ds.Tables["EI_TABLE_DEF"].Columns)
                {
                    System.Data.DataTable dt = SelectDistinct(ds.Tables["EI_TABLE_DEF"], new string[] { dc.ColumnName });
                    dt.TableName = "DISTINCT_" + dc.ColumnName;
                    ds.Tables.Add(dt);
                }
                return ds;
            }
        }

        public DataSet GetMailAddresses(string tablename)
        {
            string sql = "";
            SqlConnection connection = new SqlConnection(Connections.WinsureConnectionString);
            connection.Open();
            sql = "SELECT (SELECT EPOSTA FROM [10.81.24.180].[FSIP].[dbo].[Personel_Table] WITH (NOLOCK) WHERE SICILNO = 'FBS' + CAST(CAST (RIGHT([USER_NAME], 4) AS INT) AS VARCHAR)) MAIL " +
                    "FROM EI_MEMBER M, EI_TABLE_DEF T " +
                    "WHERE M.GROUP_ID = T.GROUP_ID AND T.TABLE_NAME = '" + tablename + "'";
            SqlDataAdapter da = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            da.Fill(ds, "EI_TABLE_MAIL");
            connection.Close();
            return ds;
        }


        public string GetTableDefinitionByName(string username, string name)
        {
            DataSet ds = GetTableDefs(username);
            return ds.Tables[0].Select("TABLE_NAME='" + name + "'")[0]["TABLE_DESCRIPTION"].ToString();
        }

        #region Distinct
        private static System.Data.DataTable SelectDistinct(System.Data.DataTable SourceTable, params string[] FieldNames)
        {
            object[] lastValues;
            System.Data.DataTable newTable;
            DataRow[] orderedRows;

            if (FieldNames == null || FieldNames.Length == 0)
                throw new ArgumentNullException("FieldNames");

            lastValues = new object[FieldNames.Length];
            newTable = new System.Data.DataTable();

            foreach (string fieldName in FieldNames)
                newTable.Columns.Add(fieldName, SourceTable.Columns[fieldName].DataType);

            orderedRows = SourceTable.Select("", string.Join(", ", FieldNames));

            foreach (DataRow row in orderedRows)
            {
                if (!fieldValuesAreEqual(lastValues, row, FieldNames))
                {
                    newTable.Rows.Add(createRowClone(row, newTable.NewRow(), FieldNames));

                    setLastValues(lastValues, row, FieldNames);
                }
            }

            return newTable;
        }

        private static bool fieldValuesAreEqual(object[] lastValues, DataRow currentRow, string[] fieldNames)
        {
            bool areEqual = true;

            for (int i = 0; i < fieldNames.Length; i++)
            {
                if (lastValues[i] == null || !lastValues[i].Equals(currentRow[fieldNames[i]]))
                {
                    areEqual = false;
                    break;
                }
            }

            return areEqual;
        }

        private static DataRow createRowClone(DataRow sourceRow, DataRow newRow, string[] fieldNames)
        {
            foreach (string field in fieldNames)
                newRow[field] = sourceRow[field];

            return newRow;
        }

        private static void setLastValues(object[] lastValues, DataRow sourceRow, string[] fieldNames)
        {
            for (int i = 0; i < fieldNames.Length; i++)
                lastValues[i] = sourceRow[fieldNames[i]];
        }
        #endregion

        public static void EmailGonder(string to, string cc, string sender, string subject, string content)
        {
            SmtpClient smtp = new SmtpClient("10.81.240.73");
            MailMessage mail = new MailMessage();

            mail.IsBodyHtml = true;
            mail.BodyEncoding = Encoding.UTF8;

            mail.From = new MailAddress("sompojapan@sompojapan.com.tr", sender);
            mail.Subject = subject;
            mail.To.Add(new MailAddress(to));
            if (cc != "")
                mail.CC.Add(new MailAddress(cc));

            mail.Body = content;

            smtp.Send(mail);
        }
        public void LogYukleme(string username, string tablename, string sonuc_file_link, string oran_file_link, DateTime gecerTarih, int taskID, int fileOrderNo, string aciklama)
        {
            string sql = "";
            SqlConnection connection = new SqlConnection(Connections.WinsureConnectionString);
            connection.Open();
            sql = "INSERT INTO EI_OP_LOG (USER_NAME, TABLE_NAME, SONUC_FILE_LINK, ORAN_FILE_LINK, OP_TYPE, REGDATE, GECER_TARIH, TASK_ID, FILE_ORDER_NO,ACIKLAMA) VALUES(@USER_NAME, @TABLE_NAME, @SONUC_FILE_LINK,@ORAN_FILE_LINK, 1,@REGDATE, @GECER_TARIH,@TASK_ID, @FILE_ORDER_NO,@ACIKLAMA)";
            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.AddWithValue("@USER_NAME", username);
            command.Parameters.AddWithValue("@TABLE_NAME", tablename);
            command.Parameters.AddWithValue("@SONUC_FILE_LINK", sonuc_file_link);
            command.Parameters.AddWithValue("@ORAN_FILE_LINK", oran_file_link);
            command.Parameters.AddWithValue("@REGDATE", DateTime.Now);
            command.Parameters.AddWithValue("@GECER_TARIH", gecerTarih);
            command.Parameters.AddWithValue("@TASK_ID", taskID);
            command.Parameters.AddWithValue("@FILE_ORDER_NO", fileOrderNo);
            command.Parameters.AddWithValue("@ACIKLAMA", aciklama);
            command.ExecuteNonQuery();
            connection.Close();
        }

        public void UpdateOpLogCanliyaAktarilanLink(int id, string sonucPath)
        {
            string sql = "";
            SqlConnection connection = new SqlConnection(Connections.WinsureConnectionString);
            connection.Open();
            sql = "UPDATE EI_OP_LOG  set LIVE_TRANSFER_LINK = @PLIVE_TRANSFER_LINK where ID = @PID";
            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.AddWithValue("@PLIVE_TRANSFER_LINK", sonucPath);
            command.Parameters.AddWithValue("@PID", id);
            command.ExecuteNonQuery();
            connection.Close();
        }


        public string LogListele(string tablename, string username, string log_bas_tarih, string log_bit_tarih)
        {
            string sql = "";
            bool kayitVarMi = false;
            SqlConnection connection = new SqlConnection(Connections.WinsureConnectionString);
            connection.Open();
            sql = "SELECT L.ID, L.USER_NAME YUKLEYEN,L.TABLE_NAME TABLO_ADI, L.SONUC_FILE_LINK YUKLENEN, L.LIVE_TRANSFER_LINK SONUC,  (CASE WHEN OP_TYPE=1 THEN 'YUKLEME' ELSE 'SILME' END) OP, L.REGDATE KAYIT_TARIHI, L.TASK_ID, L.FILE_ORDER_NO "
                + "FROM EI_OP_LOG L WITH(NOLOCK), EI_TABLE_DEF D WITH(NOLOCK), EI_GROUP G WITH(NOLOCK) ";
            if (tablename != "")
                sql = sql + "WHERE L.TABLE_NAME = @PTABLE_NAME ";
            else
                sql = sql + @"WHERE L.TABLE_NAME IN (
                    SELECT DISTINCT TABLE_NAME
                    FROM   EI_MEMBER M WITH (NOLOCK), EI_TABLE_DEF T WITH (NOLOCK)
                    WHERE M.USER_NAME=@PUSER_NAME AND T.GROUP_ID = M.GROUP_ID AND T.DELETED = 0) ";
            sql = sql + "AND D.TABLE_NAME = L.TABLE_NAME AND G.ID = D.GROUP_ID ";
            sql = sql + "AND REGDATE>=CONVERT(DATETIME, '" + log_bas_tarih + "', 104) AND REGDATE <= CONVERT(DATETIME, '" + log_bit_tarih + "', 104) + 1 ORDER BY L.REGDATE DESC, L.TABLE_NAME ASC";
            SqlCommand command = new SqlCommand(sql, connection);
            command.Parameters.AddWithValue("@PTABLE_NAME", tablename);
            command.Parameters.AddWithValue("@PUSER_NAME", username);
            SqlDataReader reader = command.ExecuteReader();
            StringBuilder txt = new StringBuilder();
            txt.Append("<table cellpadding=0 cellspacing=0 style='border:1px solid black;' align='center' class='table responsive-table-on'><thead><tr><th>Tablo</th><th>Kullanıcı</th><th>Eklenen Dosya</th><th>Sonuç Dosyası</th><th>İşlem</th><th>Kayıt Tarih</th><th>TaskID - SıraNo</th><th>Detay</th></tr></thead>");
            txt.Append("<tbody>");
            while (reader.Read())
            {
                txt.Append("<tr>");
                txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", reader["TABLO_ADI"]);
                txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", reader["YUKLEYEN"]);
                txt.AppendFormat("<td style='border:1px solid black;'><a href='{0}' target='_blank'>Dosya</a></td>", reader["YUKLENEN"]);
                if (reader["SONUC"] != null && !string.IsNullOrEmpty(reader["SONUC"].ToString()))
                {
                    txt.AppendFormat("<td style='border:1px solid black;'><a href='{0}' target='_blank'>Sonuç Dosya</a></td>", reader["SONUC"]);
                }
                else
                {
                    txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", "&nbsp;");
                }
                txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", reader["OP"]);
                txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", reader["KAYIT_TARIHI"]);
                txt.AppendFormat("<td style='border:1px solid black;'>{0}</td>", reader["TASK_ID"] + " - " + reader["FILE_ORDER_NO"]);
                txt.AppendFormat("<td style='border:1px solid black;'><a class='button full-width glossy green-gradient' href = './Default.aspx?taskid={0}' target='_blank'>Git</td>", reader["TASK_ID"]);
                txt.Append("</tr>");
                kayitVarMi = true;
            }
            if (!kayitVarMi)
            {
                txt.Append("<tr align='center'><td colspan='8'>Kayıt Bulunamadı.</td></tr>");
            }
            txt.Append("</tbody>");
            txt.Append("</table>");
            connection.Close();
            return txt.ToString();
        }

        public bool DeletePolLogByID(int id)
        {
            bool result = false;
            string sqlQuery = string.Format(@"SELECT TOP 1 [ID]
                                          ,[USER_NAME]
                                          ,[TABLE_NAME]
                                          ,[SONUC_FILE_LINK]
                                          ,[OP_TYPE]
                                          ,[REGDATE]
                                          ,[APPROVER]
                                          ,[APPROVE_DATE]
                                          ,[GECER_TARIH]
                                          ,[TASK_ID]
                                          ,[FILE_ORDER_NO]
                                          ,[SERVERNAME]
                                          ,[DATABASENAME] 
                                          ,[SCHEMANAME]
                                      FROM [WINSURE].[dbo].[EI_OP_LOG] with(nolock)
                                      where ID = @PID");

            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PID", id, SqlDbType.Int));

            System.Data.DataTable dtResult = conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);

            if (dtResult != null && dtResult.Rows.Count > 0)
            {
                try
                {
                    //File.Delete(template_path+dtResult.Rows[0]["SONUC_FILE_LINK"]); // Fiziksel olarak dosyayı sil.

                    string tableName = dtResult.Rows[0]["TABLE_NAME"].ToString();
                    string serverName = dtResult.Rows[0]["SERVERNAME"].ToString();
                    string databaseName = dtResult.Rows[0]["DATABASENAME"].ToString();
                    string schemaName = dtResult.Rows[0]["SCHEMANAME"].ToString();

                    DataBase database = new DataBase(serverName, databaseName, schemaName, tableName);
                    //string tableName = dtResult.Rows[0]["TABLE_NAME"] + "_" + DateTime.ParseExact(dtResult.Rows[0]["GECER_TARIH"].ToString(), "dd.MM.yyyy hh:mm:ss", null).ToString("yyyyMMdd");
                    DeleteRecordEITableByFileOrderID(database, Convert.ToInt32(dtResult.Rows[0]["TASK_ID"]),
                        Convert.ToInt32(dtResult.Rows[0]["FILE_ORDER_NO"])); // Temp tablo içindeki veriyi sil. Eğer 0' dan büyük fileorderno' lu yoksa tabloyu drop etsin

                    string deleteQuery = string.Format(@"delete FROM [WINSURE].[dbo].[EI_OP_LOG]
                                                      where ID = @PID");

                    int deleteResult = conTest.ExecuteNonQueryByCustomQuery(deleteQuery, param); //Oplog tablosundan ilgili kaydı silsin.

                    result = true;
                }
                catch (Exception)
                { }
            }
            return result;

        }

        public int DeleteRecordEITableByFileOrderID(DataBase database, int taskID, int fileOrderID)
        {
            string query = string.Format(@" delete [{0}].[tarifeusr].[{1}] where TASK_ID = {2} and FILE_ORDER_NO = {3} ", database.NameDatabase, database.NameTable, taskID, fileOrderID);
            var conTest = new SJF.Data.Sql.SqlManager(database.ConnectionString);
            return conTest.ExecuteNonQuery(query);
        }

        private System.Data.DataTable GetDataFromTableName(DataBase database, int taskID, int fileOrderNo)
        {
            string sqlQuery = string.Format(@"select * from {0}.{1}.{2} with(nolock) where TASK_ID = @PTASK_ID AND FILE_ORDER_NO = @PFILE_ORDER_NO", database.NameDatabase, database.TempNameSchema, database.NameTable);
            var conTest = new SJF.Data.Sql.SqlManager(database.ConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PTASK_ID", taskID, SqlDbType.Int));
            param.Add(new CommandParameter("@PFILE_ORDER_NO", fileOrderNo, SqlDbType.Int));

            return conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);
        }

        public static System.Data.DataTable GetApprover(string userName)
        {
            string sqlQuery = @"--Deleasyon Grup Listesi (Dropdownlist' in doldurulması)
                                SELECT 
	                                Cast(a.GROUP_ID as varchar)+ ' - ' +b.NAME GRUPLAR,
	                                substring(
                                    (
                                        Select ', '+c.FIRSTNAME_LASTNAME  AS [text()]
                                        From tarifeusr.EI_DELEGATION_USER c
                                        Where b.ID = c.GROUP_ID
                                        ORDER BY c.ORDER_NO
                                        For XML PATH ('')
                                    ), 2, 1000) UYELER
                                FROM tarifeusr.EI_DELEGATION_USER a with(nolock), tarifeusr.EI_DELEGATION_GROUP b with(nolock)
                                where b.ID = a.GROUP_ID
                                and a.USER_NAME = @PUSER_NAME";
            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PUSER_NAME", userName, SqlDbType.NVarChar));

            return conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);
        }

        public static System.Data.DataTable GetApproverUserList(string userName)
        {
            string sqlQuery = @"SELECT 
                                    distinct c.USER_NAME
                                FROM tarifeusr.EI_DELEGATION_USER a with(nolock), tarifeusr.EI_DELEGATION_GROUP b with(nolock), tarifeusr.EI_DELEGATION_USER c with(nolock)
                                where b.ID = a.GROUP_ID
								AND b.ID = c.GROUP_ID
                                and a.USER_NAME =@PUSER_NAME";
            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PUSER_NAME", userName, SqlDbType.NVarChar));

            return conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);
        }

        public System.Data.DataTable SeciliTabloBitmemisTasklerdeVarMi(string selectedTableName)
        {
            string sqlQuery = @"
                                SELECT 
	                                distinct his1.TASK_ID, his1.ACTION_STATUS_CHANGE_USER
                                  FROM [WINSURE].[tarifeusr].[EI_TASK] t with(nolock) 
                                  inner join [WINSURE].[dbo].[EI_OP_LOG] op with(nolock) 
                                  on op.TASK_ID = t.ID
                                  inner join [WINSURE].[tarifeusr].[EI_TASK_HISTORY] his1 with(nolock)
                                  on his1.TASK_ID = op.TASK_ID
                                  where t.ISACTIVE = 1
                                  AND op.TABLE_NAME = @PTABLE_NAME
                                  AND not exists 
                                  (
	                                SELECT 
		                                1
	                                FROM [WINSURE].[tarifeusr].[EI_TASK_HISTORY] his with(nolock)
	                                where his1.TASK_ID = his.TASK_ID
	                                and his.ACTION_STATUS_CODE in (5,7,8)
                                  )
                                  AND his1.ACTION_STATUS_CODE = 1
                                  group by his1.TASK_ID, his1.ACTION_STATUS_CHANGE_USER";
            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PTABLE_NAME", selectedTableName, SqlDbType.NVarChar));

            System.Data.DataTable dtResult = conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);

            return dtResult;
        }

        public bool WriteDataTableToExcel(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation,
            string reporType)
        {
            Microsoft.Office.Interop.Excel.Range excelCellrange;
            try
            {
                Init_Excel_Application();
                wb = EI_App.Workbooks.Add(Type.Missing);
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                sheet.Name = worksheetName.Length >= 31 ? worksheetName.Substring(0, 30) : worksheetName;

                sheet.Cells[1, 1] = reporType;
                sheet.Cells[1, 2] = "Tarih  : " + DateTime.Now.ToShortDateString();

                int rowcount = 2;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // on the first iteration we add the column headers
                        if (rowcount == 3)
                        {
                            sheet.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;
                            sheet.Cells.Font.Color = System.Drawing.Color.Black;

                        }

                        sheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        //for alternate rows
                        if (rowcount > 3)
                        {
                            if (i == dataTable.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    excelCellrange =
                                        sheet.Range[
                                            sheet.Cells[rowcount, 1], sheet.Cells[rowcount, dataTable.Columns.Count]];
                                    FormattingExcelCells(excelCellrange, "#CCCCFF", System.Drawing.Color.Black, false);
                                }

                            }
                        }

                    }

                }
                excelCellrange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rowcount, dataTable.Columns.Count]];
                excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;


                excelCellrange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, dataTable.Columns.Count]];
                FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

                Save_Excel_Application(saveAsLocation);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                excelCellrange = null;
            }
        }

        /// <summary>
        /// FUNCTION FOR FORMATTING EXCEL CELLS
        /// </summary>
        /// <param name="range"></param>
        /// <param name="HTMLcolorCode"></param>
        /// <param name="fontColor"></param>
        /// <param name="IsFontbool"></param>
        private void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

        public void ExportToExcel(System.Data.DataTable Tbl, string workSheetName, string ExcelFilePath)
        {
            try
            {
                if (Tbl == null || Tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                Init_Excel_Application();

                wb = EI_App.Workbooks.Add(Type.Missing);
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                sheet.Name = workSheetName.Length >= 31 ? workSheetName.Substring(0, 30) : workSheetName;

                // column headings
                for (int i = 0; i < Tbl.Columns.Count; i++)
                {
                    sheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;
                }

                // rows
                for (int i = 0; i < Tbl.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {
                        sheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                    }
                }

                try
                {
                    Save_Excel_Application(ExcelFilePath);

                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                        + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Close_Excel_Application(false);
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        public void DeleteTaskFile(int taskID, string tableName, int stepNo)
        {


            string sqlQuery = @"delete FROM tarifeusr.EI_TASK_FILES 
                                where TASK_ID = @PTASK_ID
                                AND STEP_NO = @PSTEP_NO
                                ";

            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PTASK_ID", taskID, SqlDbType.Int));
            param.Add(new CommandParameter("@PSTEP_NO", stepNo, SqlDbType.Int));

            if (!string.IsNullOrEmpty(tableName))
            {
                sqlQuery += " AND TABLE_NAME = @PTABLE_NAME";
                param.Add(new CommandParameter("@PTABLE_NAME", tableName, SqlDbType.VarChar));
            }


            conTest.ExecuteNonQueryByCustomQuery(sqlQuery, param);
        }

        public bool GecmisTarihliYuklemeyeIzniVarmi(string tableName)
        {
            bool result = false;
            string sqlQuery = @"SELECT top 1 [ID]
                  ,[TABLE_NAME]
                  ,[ISALLOW]
              FROM [WINSURE].[tarifeusr].[EI_GECER_TARIH_TABLE]
              where TABLE_NAME = @PTABLE_NAME";
            var conTest = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PTABLE_NAME", tableName, SqlDbType.VarChar));


            System.Data.DataTable dt = conTest.GetDataTableByCustomSqlQuery(sqlQuery, param);

            if (dt != null && dt.Rows.Count > 0)
            {
                if (Convert.ToBoolean(dt.Rows[0]["ISALLOW"]))
                {
                    result = true;
                }
            }

            return result;
        }

        private DateTime ConvertToDateTime(string strDateTime)
        {
            DateTime dtFinaldate; string sDateTime;
            try { dtFinaldate = Convert.ToDateTime(strDateTime, new CultureInfo("tr-TR")); }
            catch (Exception e)
            {
                string[] sDate = strDateTime.Split('.');
                sDateTime = sDate[1].PadLeft(2, '0') + '.' + sDate[0] + '.' + sDate[2];
                dtFinaldate = Convert.ToDateTime(sDateTime);
            }
            return dtFinaldate;
        }

        public static System.Data.DataTable GetEI_TABLE_DEF(string tableName)
        {
            string query = @"SELECT top 1 [TABLE_NAME]
                                  ,[TABLE_DESCRIPTION]
                                  ,[DELETED]
                                  ,[CATEGORY]
                                  ,[GROUP_ID]
                                  ,[MAIL_GROUP_ID]
                                  ,[TARIFE_DEG]
                                  ,[SERVERNAME]
                                  ,[DATABASENAME]
                                  ,[SCHEMANAME]
                              FROM [WINSURE].[dbo].[EI_TABLE_DEF]
                              where TABLE_NAME = @PTABLE_NAME";
            var con = new SJF.Data.Sql.SqlManager(Connections.WinsureConnectionString);

            List<CommandParameter> param = new List<CommandParameter>();
            param.Add(new CommandParameter("@PTABLE_NAME", tableName, SqlDbType.VarChar));


            System.Data.DataTable dt = con.GetDataTableByCustomSqlQuery(query, param);
            return dt;
        }
    }
}

public class SqlCommandDumper
{
    public static string GetCommandText(SqlCommand sqc)
    {
        StringBuilder sbCommandText = new StringBuilder();

        sbCommandText.AppendLine("-- BEGIN COMMAND");

        // params
        for (int i = 0; i < sqc.Parameters.Count; i++)
            logParameterToSqlBatch(sqc.Parameters[i], sbCommandText);
        sbCommandText.AppendLine("-- END PARAMS");

        // command
        if (sqc.CommandType == CommandType.StoredProcedure)
        {
            sbCommandText.Append("EXEC ");

            bool hasReturnValue = false;
            for (int i = 0; i < sqc.Parameters.Count; i++)
            {
                if (sqc.Parameters[i].Direction == ParameterDirection.ReturnValue)
                    hasReturnValue = true;
            }
            if (hasReturnValue)
            {
                sbCommandText.Append("@returnValue = ");
            }

            sbCommandText.Append(sqc.CommandText);

            bool hasPrev = false;
            for (int i = 0; i < sqc.Parameters.Count; i++)
            {
                var cParam = sqc.Parameters[i];
                if (cParam.Direction != ParameterDirection.ReturnValue)
                {
                    if (hasPrev)
                        sbCommandText.Append(", ");

                    sbCommandText.Append(cParam.ParameterName);
                    sbCommandText.Append(" = ");
                    sbCommandText.Append(cParam.ParameterName);

                    if (cParam.Direction.HasFlag(ParameterDirection.Output))
                        sbCommandText.Append(" OUTPUT");

                    hasPrev = true;
                }
            }
        }
        else
        {
            sbCommandText.AppendLine(sqc.CommandText);
        }

        sbCommandText.AppendLine("-- RESULTS");
        sbCommandText.Append("SELECT 1 as Executed");
        for (int i = 0; i < sqc.Parameters.Count; i++)
        {
            var cParam = sqc.Parameters[i];

            if (cParam.Direction == ParameterDirection.ReturnValue)
            {
                sbCommandText.Append(", @returnValue as ReturnValue");
            }
            else if (cParam.Direction.HasFlag(ParameterDirection.Output))
            {
                sbCommandText.Append(", ");
                sbCommandText.Append(cParam.ParameterName);
                sbCommandText.Append(" as [");
                sbCommandText.Append(cParam.ParameterName);
                sbCommandText.Append(']');
            }
        }
        sbCommandText.AppendLine(";");

        sbCommandText.AppendLine("-- END COMMAND");
        return sbCommandText.ToString();
    }

    private static void logParameterToSqlBatch(SqlParameter param, StringBuilder sbCommandText)
    {
        sbCommandText.Append("DECLARE ");
        if (param.Direction == ParameterDirection.ReturnValue)
        {
            sbCommandText.AppendLine("@returnValue INT;");
        }
        else
        {
            sbCommandText.Append(param.ParameterName);

            sbCommandText.Append(' ');
            if (param.SqlDbType != SqlDbType.Structured)
            {
                logParameterType(param, sbCommandText);
                sbCommandText.Append(" = ");
                logQuotedParameterValue(param.Value, sbCommandText);

                sbCommandText.AppendLine(";");
            }
            else
            {
                logStructuredParameter(param, sbCommandText);
            }
        }
    }

    private static void logStructuredParameter(SqlParameter param, StringBuilder sbCommandText)
    {
        sbCommandText.AppendLine(" {List Type};");
        var dataTable = (System.Data.DataTable)param.Value;

        for (int rowNo = 0; rowNo < dataTable.Rows.Count; rowNo++)
        {
            sbCommandText.Append("INSERT INTO ");
            sbCommandText.Append(param.ParameterName);
            sbCommandText.Append(" VALUES (");

            bool hasPrev = true;
            for (int colNo = 0; colNo < dataTable.Columns.Count; colNo++)
            {
                if (hasPrev)
                {
                    sbCommandText.Append(", ");
                }
                logQuotedParameterValue(dataTable.Rows[rowNo].ItemArray[colNo], sbCommandText);
                hasPrev = true;
            }
            sbCommandText.AppendLine(");");
        }
    }

    const string DATETIME_FORMAT_ROUNDTRIP = "o";
    private static void logQuotedParameterValue(object value, StringBuilder sbCommandText)
    {
        try
        {
            if (value == null)
            {
                sbCommandText.Append("NULL");
            }
            else
            {
                value = unboxNullable(value);

                if (value is string
                    || value is char
                    || value is char[]
                    || value is System.Xml.Linq.XElement
                    || value is System.Xml.Linq.XDocument)
                {
                    sbCommandText.Append("N'");
                    sbCommandText.Append(value.ToString().Replace("'", "''"));
                    sbCommandText.Append('\'');
                }
                else if (value is bool)
                {
                    // True -> 1, False -> 0
                    sbCommandText.Append(Convert.ToInt32(value));
                }
                else if (value is sbyte
                    || value is byte
                    || value is short
                    || value is ushort
                    || value is int
                    || value is uint
                    || value is long
                    || value is ulong
                    || value is float
                    || value is double
                    || value is decimal)
                {
                    sbCommandText.Append(value.ToString());
                }
                else if (value is DateTime)
                {
                    // SQL Server only supports ISO8601 with 3 digit precision on datetime,
                    // datetime2 (>= SQL Server 2008) parses the .net format, and will 
                    // implicitly cast down to datetime.
                    // Alternatively, use the format string "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffK"
                    // to match SQL server parsing
                    sbCommandText.Append("CAST('");
                    sbCommandText.Append(((DateTime)value).ToString(DATETIME_FORMAT_ROUNDTRIP));
                    sbCommandText.Append("' as datetime2)");
                }
                else if (value is DateTimeOffset)
                {
                    sbCommandText.Append('\'');
                    sbCommandText.Append(((DateTimeOffset)value).ToString(DATETIME_FORMAT_ROUNDTRIP));
                    sbCommandText.Append('\'');
                }
                else if (value is Guid)
                {
                    sbCommandText.Append('\'');
                    sbCommandText.Append(((Guid)value).ToString());
                    sbCommandText.Append('\'');
                }
                else if (value is byte[])
                {
                    var data = (byte[])value;
                    if (data.Length == 0)
                    {
                        sbCommandText.Append("NULL");
                    }
                    else
                    {
                        sbCommandText.Append("0x");
                        for (int i = 0; i < data.Length; i++)
                        {
                            sbCommandText.Append(data[i].ToString("h2"));
                        }
                    }
                }
                else
                {
                    sbCommandText.Append("/* UNKNOWN DATATYPE: ");
                    sbCommandText.Append(value.GetType().ToString());
                    sbCommandText.Append(" *" + "/ N'");
                    sbCommandText.Append(value.ToString());
                    sbCommandText.Append('\'');
                }
            }
        }

        catch (Exception ex)
        {
            sbCommandText.AppendLine("/* Exception occurred while converting parameter: ");
            sbCommandText.AppendLine(ex.ToString());
            sbCommandText.AppendLine("*/");
        }
    }

    private static object unboxNullable(object value)
    {
        var typeOriginal = value.GetType();
        if (typeOriginal.IsGenericType
            && typeOriginal.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            // generic value, unboxing needed
            return typeOriginal.InvokeMember("GetValueOrDefault",
                System.Reflection.BindingFlags.Public |
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.InvokeMethod,
                null, value, null);
        }
        else
        {
            return value;
        }
    }

    private static void logParameterType(SqlParameter param, StringBuilder sbCommandText)
    {
        switch (param.SqlDbType)
        {
            // variable length
            case SqlDbType.Char:
            case SqlDbType.NChar:
            case SqlDbType.Binary:
                {
                    sbCommandText.Append(param.SqlDbType.ToString().ToUpper());
                    sbCommandText.Append('(');
                    sbCommandText.Append(param.Size);
                    sbCommandText.Append(')');
                }
                break;
            case SqlDbType.VarChar:
            case SqlDbType.NVarChar:
            case SqlDbType.VarBinary:
                {
                    sbCommandText.Append(param.SqlDbType.ToString().ToUpper());
                    sbCommandText.Append("(MAX /* Specified as ");
                    sbCommandText.Append(param.Size);
                    sbCommandText.Append(" */)");
                }
                break;
            // fixed length
            case SqlDbType.Text:
            case SqlDbType.NText:
            case SqlDbType.Bit:
            case SqlDbType.TinyInt:
            case SqlDbType.SmallInt:
            case SqlDbType.Int:
            case SqlDbType.BigInt:
            case SqlDbType.SmallMoney:
            case SqlDbType.Money:
            case SqlDbType.Decimal:
            case SqlDbType.Real:
            case SqlDbType.Float:
            case SqlDbType.Date:
            case SqlDbType.DateTime:
            case SqlDbType.DateTime2:
            case SqlDbType.DateTimeOffset:
            case SqlDbType.UniqueIdentifier:
            case SqlDbType.Image:
                {
                    sbCommandText.Append(param.SqlDbType.ToString().ToUpper());
                }
                break;
            // Unknown
            case SqlDbType.Timestamp:
            default:
                {
                    sbCommandText.Append("/* UNKNOWN DATATYPE: ");
                    sbCommandText.Append(param.SqlDbType.ToString().ToUpper());
                    sbCommandText.Append(" *" + "/ ");
                    sbCommandText.Append(param.SqlDbType.ToString().ToUpper());
                }
                break;
        }
    }
}
