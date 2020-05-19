using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;
using POWord.Models;

namespace POWord.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private String connString;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "AppData\\" + "demo_poword.db";
            connString = "Data Source=" + dataPath;
        }

        public IActionResult Index()
        {
            string sql = "select * from leaveRecord order by ID DESC ";
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            StringBuilder strGrid = new StringBuilder();

            if (!dr.HasRows)
            {
                strGrid.Append("<tr class='XYDataGrid1-table-data-tr'>\r\n");
                strGrid.Append("<td colspan=4 width='100%' height='100' class='XYDataGrid1-data-cell' align='center'>对不起，暂时没有可以操作的文档。\r\n");
                strGrid.Append("</td></tr>\r\n");
            }
            else
            {
                while (dr.Read())
                {
                    strGrid.Append("<tr onmouseover='onColor(this)' onmouseout='offColor(this)' class='XYDataGrid1-table-data-tr'>\r\n");
                    strGrid.Append("<td width='7%' height='16' bgcolor='' class='XYDataGrid1-data-cell'><div align='center'><image src='images/word.gif' border='0'></image></div></td>\r\n");
                    strGrid.Append("<td width='28%' height='16' bgcolor='' class='XYDataGrid1-data-cell'><div align='left'>" + dr["Subject"].ToString() + "</div></td>\r\n");
                    strGrid.Append("<td width='20%' height='16' bgcolor='' class='XYDataGrid1-data-cell'><div align='center'><font face='宋体'>" + dr["SubmitTime"].ToString() + "</font></div></td>\r\n");
                    strGrid.Append("<td width='45%' height='16' bgcolor='' class='XYDataGrid1-data-cell'><div align='center'>\r\n");
                    strGrid.Append("<a class=OPLink href='Word/datalist?ID=" + dr["ID"].ToString() + "'  target='_blank' >数据库中字段内容</a>&nbsp;\r\n");
                    strGrid.Append("<a class=OPLink href=\"javascript:POBrowser.openWindow('Word/SubmitDataOfDoc?ID=" + dr["ID"].ToString() + "', 'width=1200px;height=800px;');\">用户填写请假条</a>&nbsp;\r\n");
                    strGrid.Append("<a class=OPLink href=\"javascript:POBrowser.openWindow('Word/GenDoc?ID=" + dr["ID"].ToString() + "', 'width=1200px;height=800px;');\">动态生成格式文档</a>&nbsp;\r\n");
                    strGrid.Append("</div></td></tr>\r\n");
                }
            }
            dr.Close();
            conn.Close();

            ViewBag.strGrid = strGrid;

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
        }
    }
}
