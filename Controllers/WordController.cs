using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;

namespace POWord.Controllers
{
    public class WordController : Controller
    {
        private string connString;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public WordController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string dataPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            dataPath = dataPath.Substring(0, dataPath.Length - 7) + "appData\\" + "demo_poword.db";
            connString = "Data Source=" + dataPath;
        }
        public IActionResult datalist()
        {

            string docID = Request.Query["ID"];

            string sql = "select * from leaveRecord where ID = " + docID;
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                //docFile = Reader["FileName"].ToString();
                ViewBag.Subject = dr["Subject"].ToString();
                ViewBag.Name = dr["Name"].ToString();
                ViewBag.Dept = dr["Dept"].ToString();
                ViewBag.Cause = dr["Cause"].ToString();
                ViewBag.Num = dr["Num"].ToString();
                ViewBag.SubmitTime = dr["SubmitTime"].ToString();
            }
            dr.Close();
            conn.Close();
            return View();
        }

         public IActionResult GenDoc()
        {


            string docID = Request.Query["ID"];

            string sql = "select * from leaveRecord where ID = " + docID;
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            doc.DisableWindowRightClick = true;
            if (dr.Read())
            {

                doc.OpenDataRegion("PO_name").Value = dr["Name"].ToString();
                doc.OpenDataRegion("PO_dept").Value = dr["Dept"].ToString();
                doc.OpenDataRegion("PO_cause").Value = dr["Cause"].ToString(); 
                doc.OpenDataRegion("PO_num").Value = dr["Num"].ToString(); ;
                doc.OpenDataRegion("PO_date").Value = dr["SubmitTime"].ToString(); ;
                doc.OpenDataRegion("PO_tip").Value = ""; ;

            }
            dr.Close();
            conn.Close();

            // 设置界面样式
            pageofficeCtrl.Caption = "动态生成格式文件";
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            // 添加自定义工具条按钮
            pageofficeCtrl.AddCustomToolButton("打印", "poPrint", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "poSetFullScreen", 4);

            ////获取数据对象
            pageofficeCtrl.SetWriter(doc);


            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/template.doc", PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }

        public IActionResult SubmitDataOfDoc()
        {

            string docID = Request.Query["ID"];

            string sql = "select * from leaveRecord where ID = " + docID;
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();

            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            PageOfficeNetCore.WordWriter.WordDocument doc = new PageOfficeNetCore.WordWriter.WordDocument();
            PageOfficeNetCore.WordWriter.DataRegion drName = doc.OpenDataRegion("PO_name");

            drName.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion drDept = doc.OpenDataRegion("PO_dept");

            drDept.Shading.BackgroundPatternColor = Color.Silver;
            //drDept.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion drCause = doc.OpenDataRegion("PO_cause");
   
            drCause.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion drNum = doc.OpenDataRegion("PO_num");
        
            drNum.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion drDate = doc.OpenDataRegion("PO_date");
            drDate.Shading.BackgroundPatternColor = Color.Pink;
            //drDate.Editing = true;
            PageOfficeNetCore.WordWriter.DataRegion drTip = doc.OpenDataRegion("PO_tip");
            drTip.Font.Italic = true;
            drTip.Value = "提示：带背景色的文字是只能通过选择设置，[]中的文字是可以录入编辑的。";



            if (dr.Read())
            {
                doc.OpenDataRegion("PO_name").Value = dr["Name"].ToString();
                doc.OpenDataRegion("PO_dept").Value = dr["Dept"].ToString();
                doc.OpenDataRegion("PO_cause").Value = dr["Cause"].ToString();
                doc.OpenDataRegion("PO_num").Value = dr["Num"].ToString(); ;
                doc.OpenDataRegion("PO_date").Value = dr["SubmitTime"].ToString(); ;
            }
            dr.Close();
            conn.Close();


            // 设置界面样式
            pageofficeCtrl.Caption = "用户填写请假条";
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;
            // 添加自定义工具条按钮
            pageofficeCtrl.AddCustomToolButton("保存", "poSave", 1);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "poSetFullScreen", 4);

            pageofficeCtrl.JsFunction_OnWordDataRegionClick = "OnWordDataRegionClick()";

            //获取数据对象
            pageofficeCtrl.SetWriter(doc);

            //设置保存页面
            pageofficeCtrl.SaveDataPage = "SaveData?ID=" + docID;
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/template.doc", PageOfficeNetCore.OpenModeType.docSubmitForm, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");

            return View();
        }

        public async Task<ActionResult> SaveData() {

            string ErrorMsg = "";
            PageOfficeNetCore.WordReader.WordDocument doc = new PageOfficeNetCore.WordReader.WordDocument(Request, Response);
            await doc.LoadAsync();
            
            string sName = doc.OpenDataRegion("PO_name").Value;
            string sDept = doc.OpenDataRegion("PO_dept").Value;
            string sCause = doc.OpenDataRegion("PO_cause").Value;
            string sNum = doc.OpenDataRegion("PO_num").Value;
            string sDate = doc.OpenDataRegion("PO_date").Value;

            if (sName == "")
            {
                ErrorMsg = ErrorMsg + "<li>申请人</li>";
            }
            if (sDept == "")
            {
                ErrorMsg = ErrorMsg + "<li>部门名称</li>";
            }
            if (sCause == "")
            {
                ErrorMsg = ErrorMsg + "<li>请假原因</li>";
            }
            if (sDate == "")
            {
                ErrorMsg = ErrorMsg + "<li>日期</li>";
            }

            try
            {
                if (sNum != "")
                {
                    if (Int64.Parse(sNum) < 0)
                    {
                        ErrorMsg = ErrorMsg + "<li>请假天数不能是负数</li>";
                    }
                }
                else
                {
                    ErrorMsg = ErrorMsg + "<li>请假天数</li>";
                }
            }
            catch (Exception Ex)
            {
                ErrorMsg = ErrorMsg + "<li><font color=red>注意：</font>请假天数必须是数字</li>";
            }

            string id = Request.Query["ID"];
            string sql = "";
            if (ErrorMsg == "")
            {
                sql = "update leaveRecord set Name='" + sName
                + "', Dept='" + sDept + "', Cause='" + sCause
                + "', Num=" + sNum + ", SubmitTime='" + sDate
                + "' where  ID=" + id;
                SqliteConnection conn = new SqliteConnection(connString);
                conn.Open();
                SqliteCommand cmd = new SqliteCommand(sql, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            else
            {
                await Response.Body.WriteAsync(Encoding.GetEncoding("GB2312").GetBytes(ErrorMsg));
                doc.ShowPage(578, 380);
            }
            doc.Close();
            return Content("OK");
        }

    }
}