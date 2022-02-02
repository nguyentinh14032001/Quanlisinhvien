using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using Newtonsoft.Json;
using PagedList;
using projectCK.Models;
using Rotativa;
using Rotativa.MVC;

namespace projectCK.Controllers
{
    public class NganhsController : Controller
    {
        private data db = new data();

        // GET: Nganhs
        public ActionResult Index()
        {
            ViewBag.MaKhoa = new SelectList(db.Khoas, "MaKhoa", "TenKhoa");
            var list = db.Nganhs.Include(n => n.Khoa).Where(m => m.TrangThai != 0).ToList();
            return View(list.ToList());
        }
        public ActionResult sv_Nganh_Index()
        {
            int count = db.Nganhs.Count(m => m.TrangThai != 0);
            ViewBag.SoLuong = count;
            var nganhs = db.Nganhs.Include(n => n.Khoa);
            return View(nganhs.ToList());
        }
        public ActionResult gv_Nganh_Index()
        {
            int count = db.Nganhs.Count(m => m.TrangThai != 0);
            ViewBag.SoLuong = count;
            var nganhs = db.Nganhs.Include(n => n.Khoa);
            return View(nganhs.ToList());
        }

        // GET: Nganhs/Details/5

        [HttpGet]
        public JsonResult Getdata2()
        {
            var lists = db.Khoas.ToList();
            var a = JsonConvert.SerializeObject(lists);
            return Json(a, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public void AddData(string mn, string tn, string mk, string stc)
        {

            string manganh = mn;
            string tennganh = tn;
            string makhoa = mk;
            string tsotinchi = stc;
                var data = db.Nganhs.Where(s => s.MaNganh.Equals(manganh)).ToList();
                if (data.Count > 0)
                {
                    Response.Write("Trùng mã ngành");
                }
                else
                {
                    Nganh ng = new Nganh();
                    ng.MaNganh = manganh;
                    ng.TenNganh = tennganh;
                    ng.MaKhoa = makhoa;
                    ng.TongSTC = tsotinchi;
                    db.Nganhs.Add(ng);
                    db.SaveChanges();
                    Response.Write("Thêm thành công");

                }
            }
        [HttpPost]
        public void EditData(string mn, string tn, string mk, string stc)
        {

            string manganh = mn;
            string tennganh = tn;
            string makhoa = mk;
            string tsotinchi = stc;
            try
            {
                Nganh ng = new Nganh();
                ng = db.Nganhs.Where(s => s.MaNganh == manganh).SingleOrDefault();
                ng.TenNganh = tennganh;
                ng.MaKhoa = makhoa;
                ng.TongSTC = tsotinchi;
                db.SaveChanges();
                Response.Write("Sửa thành công");

            }
            catch 
            {

            }
        }
        [HttpGet]
        public JsonResult Getdata1(string id)
        {
            var list = (from l in db.Nganhs
                        from c in db.Khoas
                        where l.MaKhoa == c.MaKhoa && l.MaNganh.Equals(id)
                        select new
                        {
                            l.MaNganh,
                            l.TenNganh,
                            l.MaKhoa,
                            c.TenKhoa,
                            l.TongSTC,
                        });
            var a = JsonConvert.SerializeObject(list);
            return Json(a, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult Getdata3(string id)
        {
            
            var list = (from l in db.Nganhs
                        from c in db.Khoas
                        where l.MaKhoa == c.MaKhoa && l.TenNganh.Contains(id)
                        select new
                        {
                           l.MaNganh,
                            l.TenNganh,
                            l.MaKhoa,
                            c.TenKhoa,
                            l.TongSTC,
                        });
            var a = JsonConvert.SerializeObject(list);
            return Json(a, JsonRequestBehavior.AllowGet);
        }
        public ActionResult DelTrash(string id)
        {
            Nganh nganh = db.Nganhs.Find(id);
            nganh.TrangThai = 0;
            //Cập nhật lại
            db.Entry(nganh).State = EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        public ActionResult Trash()
        {
            int count = db.Nganhs.Count(m => m.TrangThai == 0);
            ViewBag.SoLuong = count;
            var list = db.Nganhs.Where(m => m.TrangThai == 0).ToList();
            return View("Trash", list);
        }
        public ActionResult ReTrash(string id)
        {
            Nganh nganh = db.Nganhs.Find(id);
            nganh.TrangThai = 1;
            db.Entry(nganh).State = EntityState.Modified;
            db.SaveChanges();
            return RedirectToAction("Trash", "Nganhs");
        }
        [HttpPost]
        public ActionResult XuatFile(string file)
        {
            if (file == "1")
            {
                var gv = new GridView();
                gv.DataSource = db.Nganhs.Where(m => m.TrangThai != 0).ToList();
                gv.DataBind();

                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=nganh.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                objHtmlTextWriter.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
                gv.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
                return View();
            }
            else
            {
                if (file == "2")
                {
                    //sử dụng code https://techfunda.com/howto/309/export-data-into-ms-word-from-mvc
                    // get the data from database
                    var data = db.Nganhs.Where(m => m.TrangThai != 0).ToList();
                    // instantiate the GridView control from System.Web.UI.WebControls namespace
                    // set the data source
                    GridView gridview = new GridView();
                    gridview.DataSource = data;
                    gridview.DataBind();

                    // Clear all the content from the current response
                    Response.ClearContent();
                    Response.Buffer = true;
                    // set the header
                    Response.AddHeader("content-disposition", "attachment;filename = nganh.doc");
                    Response.ContentType = "application/ms-word";
                    Response.Charset = "";
                    // create HtmlTextWriter object with StringWriter
                    using (StringWriter sw = new StringWriter())
                    {
                        using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                        {
                            htw.WriteLine("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
                            // render the GridView to the HtmlTextWriter
                            gridview.RenderControl(htw);
                            // Output the GridView content saved into StringWriter
                            Response.Output.Write(sw.ToString());
                            Response.Flush();
                            Response.End();
                        }
                    }
                    return View();
                }
                else
                {
                    // sử dụng thư viện https://www.c-sharpcorner.com/article/export-pdf-from-html-in-mvc-net4/
                    var list = db.Nganhs.Where(m => m.TrangThai != 0).ToList();
                    return new PartialViewAsPdf("PrintPDF", list)
                    {
                        FileName = "Nganh.pdf"
                    };
                }
            }
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase file)
        {
            string filePath = string.Empty;
            if (file != null)
            {
                string path = Server.MapPath("~/Content/Uploads/");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                filePath = path + Path.GetFileName(file.FileName);
                string extension = Path.GetExtension(file.FileName);
                file.SaveAs(filePath);
                string conString = string.Empty;

                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        //Provider=Microsoft.Jet.OLEDB.4.0
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Microsoft.ACE.OLEDB.12.0" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                    case ".xlsx": //Excel 07 and above.
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES'";
                        break;
                }

                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            //Get the name of First Sheet.
                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            //Read Data from First Sheet.
                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(dt);
                            connExcel.Close();
                        }
                    }
                }
                conString = @"Data Source=TINH\SQLEXPRESS;Initial Catalog=qlsinhvien;Integrated Security=True";
                using (SqlConnection con = new SqlConnection(conString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        //Set the database table name.
                        sqlBulkCopy.DestinationTableName = "Nganh";

                        // Map the Excel columns with that of the database table, this is optional but good if you do
                        // 
                        sqlBulkCopy.ColumnMappings.Add("MaNganh", "MaNganh");
                        sqlBulkCopy.ColumnMappings.Add("TenNganh", "TenNganh");
                        sqlBulkCopy.ColumnMappings.Add("MaKhoa", "MaKhoa");
                        sqlBulkCopy.ColumnMappings.Add("TongSTC", "TongSTC");
                        sqlBulkCopy.ColumnMappings.Add("TrangThai", "TrangThai");
                        con.Open();
                        sqlBulkCopy.WriteToServer(dt);
                        con.Close();
                    }
                }
            }
            //if the code reach here means everthing goes fine and excel data is imported into database
            ViewBag.Success = "File Imported and excel data saved into database";

            // return View(db.Khoas.ToList());
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult TimDrop(string Makhoa)
        {
                     
            ViewBag.MaKhoa = new SelectList(db.Khoas, "MaKhoa", "TenKhoa");
            var list = db.Nganhs.Where(s => s.TrangThai != 0).ToList();
            if (Makhoa == "") {
                int count = db.Nganhs.Count(m => m.TrangThai != 0);
                ViewBag.SoLuong = count;
                list = db.Nganhs.Where(s => s.TrangThai != 0).ToList();
            }
            else
            {
                int count = db.Nganhs.Count(m => m.MaKhoa.Equals(Makhoa) && m.TrangThai != 0);
                ViewBag.SoLuong = count;
                list = db.Nganhs.Where(m => m.MaKhoa.Equals(Makhoa) && m.TrangThai != 0).ToList();
            }

            return View("Index", list.ToPagedList(1,5));
        }
        [HttpGet]
        public ActionResult Index(int? page)
        {
            
            ViewBag.MaKhoa = new SelectList(db.Khoas, "MaKhoa", "TenKhoa");
            int count = db.Nganhs.Count(m => m.TrangThai != 0);
            ViewBag.SoLuong = count;
         
            var list = db.Nganhs.Include(n => n.Khoa).Where(m => m.TrangThai != 0).ToList();

            // 1. Tham số int? dùng để thể hiện null và kiểu int
            // page có thể có giá trị là null và kiểu int.

            // 2. Nếu page = null thì đặt lại là 1.
            if (page == null) page = 1;

            // 3. Tạo truy vấn, lưu ý phải sắp xếp theo trường nào đó, ví dụ OrderBy
            // theo LinkID mới có thể phân trang.


            // 4. Tạo kích thước trang (pageSize) hay là số Link hiển thị trên 1 trang
            int pageSize = 5;

            // 4.1 Toán tử ?? trong C# mô tả nếu page khác null thì lấy giá trị page, còn
            // nếu page = null thì lấy giá trị 1 cho biến pageNumber.
            int pageNumber = (page ?? 1);

            // 5. Trả về các Link được phân trang theo kích thước và số trang.
            

          
            return View(list.ToPagedList(pageNumber, pageSize));
           
        }
        public void Delete(string manganh)
        {
            try
            {
                Nganh nganh= db.Nganhs.Find(manganh);
                db.Nganhs.Remove(nganh);
                db.SaveChanges();
                Response.Write("Xóa thành công");
            }
            catch
            {
                Response.Write("Xóa thất bại");
            }
        }



        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    
      


    }

}
