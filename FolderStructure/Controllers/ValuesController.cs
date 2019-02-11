using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.IO;
using DSOFile;

namespace FolderStructure.Controllers
{
    public class ValuesController : ApiController
    {
        public DataTable dt = null;

        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/values
        [HttpPost]
        public IHttpActionResult GetFolderContents(string path)
        {
            dt = new DataTable();
            dt.Columns.Add("title");
            dt.Columns.Add("type");
            dt.Columns.Add("path");
            dt.Columns.Add("createdBy");
            dt.Columns.Add("lastModified");
            string basePath = "D:\\";
            System.Collections.IEnumerable subDirectories;
            //System.Collections.IEnumerable files;
            //string[] files;
            List<string> files;
            try
            {
                if (path != null)
                {
                    basePath = path;
                }
                //subDirectories = Directory.GetDirectories(basePath);
                subDirectories = Directory.GetDirectories(basePath);
                files = Directory.GetFiles(basePath).ToList();

                DSOFile.OleDocumentProperties file1 = new DSOFile.OleDocumentProperties();
                foreach (var item in files)
                {
                    


                    DataRow dr = dt.NewRow();
                    string[] titleTokens = item.ToString().Split(new[] { '\\' }, StringSplitOptions.None);
                    string title = titleTokens[titleTokens.Length - 1];

                    dr["title"] = title;
                    dr["type"] = "FILE";
                    dr["path"] = item;
                    dr["lastModified"] = Directory.GetLastWriteTime(item.ToString());


                    // Get Custom Property
                    file1.Open(@item, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);

                    foreach (DSOFile.CustomProperty property in file1.CustomProperties)
                    {
                        if (property.Name == "createdBy")
                        {
                            //Property exists
                            dr["createdBy"] = property.get_Value();
                        }
                    }
                    file1.Close();


                    dt.Rows.Add(dr);

                    
                }
                foreach (var item in subDirectories)
                {
                    DataRow dr = dt.NewRow();
                    string[] titleTokens = item.ToString().Split(new[] { '\\' }, StringSplitOptions.None);
                    string title = titleTokens[titleTokens.Length - 1];

                    dr["title"] = title;
                    dr["type"] = "DIR";
                    dr["path"] = item;
                    dr["lastModified"] = Directory.GetLastWriteTime(item.ToString());

                    if (!IsSystemDir(item.ToString()))
                    {
                        // Get Custom Property
                        file1.Open(@item.ToString(), false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);

                        foreach (DSOFile.CustomProperty property in file1.CustomProperties)
                        {
                            if (property.Name == "createdBy")
                            {
                                //Property exists
                                dr["createdBy"] = property.get_Value();
                            }
                        }
                        file1.Close();
                    }

                    dt.Rows.Add(dr);
                }

                return Ok(dt);
            }
            catch (Exception e)
            {
                throw;
            }
        }
        private static bool IsSystemDir(string dir)
        {
            if (dir.EndsWith("System Volume Information")) return true;
            if (dir.Contains("$RECYCLE.BIN")) return true;
            return false;
        }
        [HttpPost]
        public IHttpActionResult createNewFolder()
        {
            
            try
            {
                string path = HttpContext.Current.Request.Form["path"];
                string directoryName = HttpContext.Current.Request.Form["directoryName"];
                string createdBy = HttpContext.Current.Request.Form["createdBy"];

                if (!path.EndsWith("\\"))
                {
                    path = path + "\\";
                }
                Directory.CreateDirectory(path + directoryName);

                // Add custom property
                DSOFile.OleDocumentProperties file1 = new DSOFile.OleDocumentProperties();
                file1.Open(path + directoryName, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);
                file1.CustomProperties.Add("createdBy", createdBy);
                file1.Close(true);
                
                return Ok();
            }
            catch (Exception e)
            {
                throw;
            }
        }
        [HttpPost]
        public IHttpActionResult uploadFile()
        {
            try
            {
                if (HttpContext.Current.Request.Files.AllKeys.Any())
                {
                    //string fileName = HttpContext.Current.Request.Form["fileName"];

                    // Upload file
                    string path = HttpContext.Current.Request.Form["path"];
                    if (!path.EndsWith("\\"))
                    {
                        path = path + "\\";
                    }
                    var file = HttpContext.Current.Request.Files[0];
                    file.SaveAs(path + file.FileName);

                    // Add custom property
                    DSOFile.OleDocumentProperties file1 = new DSOFile.OleDocumentProperties();
                    file1.Open(path + file.FileName, false, DSOFile.dsoFileOpenOptions.dsoOptionDefault);
                    object createdBy = HttpContext.Current.Request.Form["createdBy"];
                    file1.CustomProperties.Add("createdBy", createdBy);
                    file1.Close(true);

                    return Ok();
                }
                return BadRequest();
            }
            catch (Exception e)
            {
                throw;
            }
        }
        //// PUT api/values/5
        //public void Put(int id, [FromBody]string value)
        //{
        //}

        //// DELETE api/values/5
        //public void Delete(int id)
        //{
        //}
    }
}
