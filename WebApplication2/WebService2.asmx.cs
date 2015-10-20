using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Web.Script.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Microsoft.Office.Interop.Word;
using System.Net;


namespace WebApplication2
{
    /// <summary>
    /// WebService2 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class WebService2 : System.Web.Services.WebService
    {

        public class hahabaseobject
        {
            public string title;
            public string description;
            public ArrayList Source = new ArrayList();
            public string othercontext;
            public string Implement;
            public string Priority;
            public string Contribution;
            public string Category;
            public string Allocation;
            public string sources;

        }
        public class step123
        {
            public int num;
            public string actions;
            public string expected_result;
            public string indata;
            public string test_step;

        }
        public class tctable
        {
            public string tag;
            public string description;
            public string test_item;
            public string test_method;
            public string pre_condition;
            public string result;
            public string comment;
            public ArrayList steps = new ArrayList();
            public ArrayList source = new ArrayList();

            public string input;
            public string exec_step;
            public string exp_step;
            //public ArrayList actions = new ArrayList();
            //public ArrayList epr = new ArrayList();
        }
        public class finaljson
        {
            public ArrayList finalstrings = new ArrayList();

        }
        public class JsonTools
        {
            // 从一个对象信息生成Json串
            public static string ObjectToJson(object obj)
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                MemoryStream stream = new MemoryStream();
                serializer.WriteObject(stream, obj);
                byte[] dataBytes = new byte[stream.Length];
                stream.Position = 0;
                stream.Read(dataBytes, 0, (int)stream.Length);
                return Encoding.UTF8.GetString(dataBytes);
            }
            // 从一个Json串生成对象信息
            public static object JsonToObject(string jsonString, object obj)
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                MemoryStream mStream = new MemoryStream(Encoding.UTF8.GetBytes(jsonString));
                return serializer.ReadObject(mStream);
            }
        }
        public string swfnanme;
        private string OfficeToPdf(string OfficePath, string OfficeName, string destPath)
        {
            string fullPathName = OfficePath + OfficeName;//包含 路径 的全称
            FileInfo fi1 = new FileInfo(fullPathName);
            fi1.Attributes = ~FileAttributes.ReadOnly;
            string fileNameWithoutEx = System.IO.Path.GetFileNameWithoutExtension(OfficeName);//不包含路径，不包含扩展名
            string extendName = System.IO.Path.GetExtension(OfficeName).ToLower();//文件扩展名
            string saveName = destPath + fileNameWithoutEx + ".pdf";
            string returnValue = fileNameWithoutEx + ".pdf";
            Util.WordToPDF(fullPathName, saveName);
            return returnValue;
        }
        private string PdfToSwf(string pdf2swfPath, string PdfPath, string PdfName, string destPath)
        {
            string fullPathName = PdfPath + PdfName;//包含 路径 的全称
            string fileNameWithoutEx = System.IO.Path.GetFileNameWithoutExtension(PdfName);//不包含路径，不包含扩展名
            string extendName = System.IO.Path.GetExtension(PdfName).ToLower();//文件扩展名
            string saveName = destPath + fileNameWithoutEx + ".swf";
            string returnValue = fileNameWithoutEx + ".swf"; ;
            Util.PDFToSWF(pdf2swfPath, fullPathName, saveName);
            return returnValue;
        }
        public string showwordfiles(string filename)
        {
            string pdf2swfToolPath = System.Web.HttpContext.Current.Server.MapPath("~/FlexPaper/pdf2swf.exe");
            string OfficeFilePath = "D://pdf/office/";
            string PdfFilePath = "D://pdf/pdf/";
            string SWFFilePath = "D://pdf/swf/";
            string SwfFileName = String.Empty;
            string UploadFileName = System.IO.Path.GetFileNameWithoutExtension(filename) + ".doc";
            string UploadFileType = System.IO.Path.GetExtension(UploadFileName).ToLower();
            string UploadFileNameFileFullName = String.Empty;
            UploadFileNameFileFullName = OfficeFilePath + UploadFileName;
            File.Copy(filename, UploadFileNameFileFullName);
            string PdfFileName = OfficeToPdf(OfficeFilePath, UploadFileName, PdfFilePath);
            SwfFileName = PdfToSwf(pdf2swfToolPath, PdfFilePath, PdfFileName, SWFFilePath);
            return SWFFilePath;
        }
        public string delet_tables(string filename)
        {
            _Application appdelet_tables = new Microsoft.Office.Interop.Word.Application();
            _Document docdelet_tables;
            object fileName = filename;
            object unknow = System.Type.Missing;
            docdelet_tables = appdelet_tables.Documents.Open(ref fileName,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
            int table_num = docdelet_tables.Tables.Count;
            try
            {
                for (int i = 1; i <= table_num; i++)
                {
                    docdelet_tables.Tables[i].Delete();
                }
            }
            catch { }
            docdelet_tables.Close(ref unknow, ref unknow, ref unknow);

            appdelet_tables.Quit(ref unknow, ref unknow, ref unknow);
            return null;
        }
       
        public void thread1()
        {
            string temp = null;
            int start = 1;
            int end = pcount ;

            Dictionary<string, string> map1=null;
            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) {
                     map1 = new Dictionary<string, string>(); map1.Add("tag", temp);
                }
                if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                       map1.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map1);
                }
                else
                {



                }


                start++;







            }//while
        }//thread1


        public void thread2()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount / 8 + 1;
            int end = pcount * 2 / 8;
           
            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map2 = new Dictionary<string, string>(); map2.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map2.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map2);
                }
                else
                {



                }


                start++;







            }//while
            }
         
        public void thread3()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 2 / 8 + 1;
            int end = pcount * 3 / 8;

            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map3 = new Dictionary<string, string>(); map3.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map3.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map3);
                }
                else
                {



                }


                start++;







            }//while
        }
        public void thread4()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 3 / 8 + 1;
            int end = pcount * 4 / 8;

            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map4 = new Dictionary<string, string>(); map4.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map4.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map4);
                }
                else
                {



                }


                start++;







            }//while
        }
        public void thread5()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 4 / 8 + 1;
            int end = pcount * 5 / 8;

            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map5 = new Dictionary<string, string>(); map5.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map5.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map5);
                }
                else
                {



                }


                start++;







            }//while
        }
        public void thread6()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 5 / 8 + 1;
            int end = pcount * 6 / 8;

            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map6 = new Dictionary<string, string>(); map6.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map6.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map6);
                }
                else
                {



                }


                start++;







            }//while
        }
        public void thread7()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 6 / 8 + 1;
            int end = pcount * 7 / 8;

            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map7 = new Dictionary<string, string>(); map7.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map7.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map7);
                }
                else
                {



                }


                start++;







            }//while
        }
        public void thread8()
        {
            int arraycount = 0;
            string temp = null;
            string tempf = null;
            int start = pcount * 7 / 8 + 1;
            int end = pcount;
           
            while (start <= end)
            {

                temp = doc.Paragraphs[start].Range.Text.Trim();
                if (Regex.Matches(temp, @"(^\[TSP-.*?-\d*?\]$)", RegexOptions.IgnoreCase).Count > 0) { map8 = new Dictionary<string, string>(); map8.Add("tag", temp); }
                else if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                {


                    //大小写的问题尚未解决呢

                    Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                    // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                    if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                    {

                        map8.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                    }

                }
                else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                {
                    aaa.finalstrings.Add(map8);
                }
                else
                {



                }


                start++;







            }//while

           
        }//多线程并行处理
        public _Document doc;
        public _Document doc1;
        public _Document doc2;
        public _Document doc3;
        public _Document doc4;
        public _Document doc5;
        public _Document doc6;
        public _Document doc7;
        public string in_column;
        Dictionary<string, string> map2;
        Dictionary<string, string> map3;
        Dictionary<string, string> map4;
        Dictionary<string, string> map5;
        Dictionary<string, string> map6;
        Dictionary<string, string> map7;
        Dictionary<string, string> map8;
        
        public string pattern1;
        public string pattern2;
        public string pattern3;
        public string pattern4;
        public string type1;
        public string type2;
        public int pcount;
        string[] columns;
        public string string1;
        public string string2;
        public string string3;
        public string string4;
        public string string5;
        public string string6;
        public string string7;
        public string string8;
        public int threadsig1 = 0;
        public int threadsig2 = 0;
        public string root = "E:\\files\\";//文件保存的路径;
        public finaljson aaa = new finaljson();
        [WebMethod]
        public bool downfile(string url)
        {
            try
            {
                //return false;

                int poseuqlurl = url.IndexOf('=');
                string url1;
                url1 = url.Substring(poseuqlurl + 1, url.Length - poseuqlurl - 1);
                Uri u = new Uri(url1);
                string filename = DateTime.Now.ToString() + ".doc";
                string LocalPath = "D:\\" + filename;
                HttpWebRequest mRequest = (HttpWebRequest)WebRequest.Create(u);
                mRequest.Method = "GET";
                mRequest.ContentType = "application/x-www-form-urlencoded";
                HttpWebResponse wr = (HttpWebResponse)mRequest.GetResponse();
                Stream sIn = wr.GetResponseStream();
                FileStream fs = new FileStream(LocalPath, FileMode.Create, FileAccess.Write);
                //BinaryWriter brnew = new BinaryWriter(fs);
                //brnew.Write(bytContent, 0, bytContent.Length);
                byte[] bytes = new byte[4096];

                int start = 0;

                int length;

                while ((length = sIn.Read(bytes, 0, 4096)) > 0)
                {

                    fs.Write(bytes, 0, length);

                    start += length;

                }
                sIn.Close();
                wr.Close();
                fs.Close();
                return true;
            }
            catch { return false; }
        }
       
        [WebMethod(Description = "readtc_void")]
        public void readtc()
        {
            Microsoft.Office.Interop.Word.Table nowTable;
            //建模已经完成啦嘎嘎
            int end=doc.Tables.Count;

            for (int tablePos = 1; tablePos <= end; tablePos++)
            {
                 nowTable= doc.Tables[tablePos];
                 Regex  tsp_mark = new Regex(@"^(\[TSP-.*?-\d*?\]).*?");
                 if (!tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Success)
                 {
                    // aaa.finalstrings.Add(nowTable.Cell(1, 2).Range.Text.Trim());
                     continue;

                 }
                // aaa.finalstrings.Add(columns.ToString());
                 Dictionary<string,object> h = new Dictionary<string,object>();
                 h.Add("tag",tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Groups[1].Value);
                 for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
                 {  //还要把中文给去掉,以及纵向合并的单元格

                     
                     string flag = Regex.Match(nowTable.Rows[rowPos].Cells[1].Range.Text.Trim().ToLower().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase).Groups[1].Value;//,@"\w",RegexOptions.IgnoreCase).Groups[0].Value.Trim();
                     string[] flags=flag.Split(new char[] {' '});
                     flag=string.Join(" ",flags);
                  //   aaa.finalstrings.Add(flag);
                     bool mark = false ;
                     foreach (string a in columns)
                     {
                        if(a.ToLower().Trim()==flag){mark=true;break;}
                         
                     }

                     //aaa.finalstrings.Add(flag);
                   
                     if(mark&&nowTable.Rows[rowPos].Cells.Count==2){
                       //  aaa.finalstrings.Add(flag + "haha here" + nowTable.Rows[rowPos].Cells[1].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                         h.Add(flag,nowTable.Rows[rowPos].Cells[2].Range.Text.Trim().Replace("\r","").Replace("\u0007",""));


                     }else if(mark&&nowTable.Rows[rowPos].Cells.Count>2){
                         //处理tc_test的情况了的
                         //收集列名字
                      //   aaa.finalstrings.Add(flag + "tc_step here" + in_column.Trim());
                          int num=nowTable.Rows[rowPos].Cells.Count;
                         string [] column_name=new string[num+1];
                         for(int i=1;i<=nowTable.Rows[rowPos].Cells.Count;i++){

                             Match temp = Regex.Match(nowTable.Rows[rowPos].Cells[i].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase);
                             column_name[i] = temp.Groups[1].Value.ToLower();// +temp[1].Value + nowTable.Rows[rowPos].Cells[i].Range.Text.Trim();
                             
                         }
                             column_name[1]="num";
                            h["test steps"] = null;
                      //往下收集列即可
                           int k = 0;
                         for(int start=rowPos+1;start<nowTable.Rows.Count;start++){

                             if(nowTable.Rows[start].Cells.Count!=num){
                                rowPos=start-1;break;
                             }
                           Dictionary<string,string>   tc_step = new Dictionary<string, string>();
                              int j;
                             for(j=1;j<=num;j++){

                                tc_step.Add(column_name[j], nowTable.Rows[start].Cells[j].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                

                            }
                             h["test steps"]+=(new JavaScriptSerializer().Serialize(tc_step))+",";
                              

                         }//处理tc_step
                          h["test steps"]="["+h["test steps"].ToString().Substring(0, h["test steps"].ToString().Length-1)+"]";
                     }//else if 
                 }//for row

                 aaa.finalstrings.Add(h);

                }//for tables

                 
            }

 









       
        [WebMethod(Description = "readtc")]
        public  string  resolve(string column, string type, string doc_url = "http://127.0.0.1/casco-api/public/files/tcs.doc")
        {
            String LocalPath=null;
            try
            {   
                LocalPath = root+ doc_url.Substring(doc_url.LastIndexOf('/'));
                Uri u = new Uri(doc_url);
                
                HttpWebRequest mRequest = (HttpWebRequest)WebRequest.Create(u);
                mRequest.Method = "GET";
                mRequest.ContentType = "application/x-www-form-urlencoded";
                HttpWebResponse wr = (HttpWebResponse)mRequest.GetResponse();
                Stream sIn = wr.GetResponseStream();
                FileStream fs = new FileStream(LocalPath, FileMode.Create, FileAccess.Write);
                byte[] bytes = new byte[4096];
                int start = 0;
                int length;
                while ((length = sIn.Read(bytes, 0, 4096)) > 0)
                {
                    fs.Write(bytes, 0, length);
                    start += length;
                }
                sIn.Close();
                wr.Close();
                fs.Close();
            }
            catch { }
            if (!File.Exists(LocalPath)) { return "{code:0,msg:保存文件失败！=>"+doc_url+"}";  }
          
            in_column = column;
            //WORD  中数据都规整为一个空格隔开来
            columns = column.Split(',');
           
                _Application app = new Microsoft.Office.Interop.Word.Application();
         /*       _Application app1= new Microsoft.Office.Interop.Word.Application();
                _Application app2= new Microsoft.Office.Interop.Word.Application();
                _Application app3= new Microsoft.Office.Interop.Word.Application();
                _Application app4= new Microsoft.Office.Interop.Word.Application();
                _Application app5= new Microsoft.Office.Interop.Word.Application();
                _Application app6= new Microsoft.Office.Interop.Word.Application();
                _Application app7= new Microsoft.Office.Interop.Word.Application();
          */
             
                object fileName =LocalPath;

                object unknow = System.Type.Missing;
               
               //目前一个线程再跑
                doc = app.Documents.Open(ref fileName,
                    ref unknow, true, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
         /*       doc1 = app1.Documents.Open(ref fileName,
                   ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                   ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                   ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc2 = app2.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc3 = app3.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc4 = app4.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc5 = app5.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc6 = app6.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
                doc7 = app7.Documents.Open(ref fileName,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                    ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
          * */
                pcount = doc.Paragraphs.Count;//count the paragraphs
                int jjj = pcount;

                if (type == "rs")
                {
                    //然后完成对文档的解析

             
                var t1 = new System.Threading.Tasks.Task(() => thread1());
             /*   var t2 = new System.Threading.Tasks.Task(() => thread1());
                var t3 = new System.Threading.Tasks.Task(() => thread1());
                var t4 = new System.Threading.Tasks.Task(() => thread1());
                var t5 = new System.Threading.Tasks.Task(() => thread1());
                var t6 = new System.Threading.Tasks.Task(() => thread1());
                var t7 = new System.Threading.Tasks.Task(() => thread1());
                var t8 = new System.Threading.Tasks.Task(() => thread1());
                t1.Start();
                t2.Start();
                t3.Start();
                t4.Start();
                t5.Start();
                t6.Start();
                t7.Start();
                t8.Start();
              * */
                t1.Start();
                System.Threading.Tasks.Task.WaitAll(t1);//, t2, t3, t4, t5, t6, t7, t8);
                }//rs
                else if (type == "tc")
                {

                    var t1 = new System.Threading.Tasks.Task(() => readtc());
                    t1.Start();
                    System.Threading.Tasks.Task.WaitAll(t1);


                }

                doc.Close(ref unknow, ref unknow, ref unknow);

              /*  doc1.Close(ref unknow, ref unknow, ref unknow);

                doc2.Close(ref unknow, ref unknow, ref unknow);

                doc3.Close(ref unknow, ref unknow, ref unknow);

                doc4.Close(ref unknow, ref unknow, ref unknow);

                doc5.Close(ref unknow, ref unknow, ref unknow);

                doc6.Close(ref unknow, ref unknow, ref unknow);

                doc7.Close(ref unknow, ref unknow, ref unknow);
*/
                app.Quit(ref unknow, ref unknow, ref unknow);
           /*     app1.Quit(ref unknow, ref unknow, ref unknow);
                app2.Quit(ref unknow, ref unknow, ref unknow);
                app3.Quit(ref unknow, ref unknow, ref unknow);
                app4.Quit(ref unknow, ref unknow, ref unknow);
                app5.Quit(ref unknow, ref unknow, ref unknow);
                app6.Quit(ref unknow, ref unknow, ref unknow);
                app7.Quit(ref unknow, ref unknow, ref unknow);
            */
             
                
                var json = new JavaScriptSerializer().Serialize(aaa.finalstrings);
                return  json;
                Context.Response.ContentType = "text/json";
                Context.Response.Write(json);
                Context.Response.End();
               
                GC.Collect();
                GC.Collect();

        }




        [WebMethod(Description = "readtitles")]
        public string readtitles(string filename)
        {
            _Application app = new Microsoft.Office.Interop.Word.Application();
            _Document doc;

            object fileName = filename;
            object unknow = System.Type.Missing;
            doc = app.Documents.Open(ref fileName,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                           ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);//input a doc
            object pcount = doc.Paragraphs.Count;//count the paragraphs
            object trydocfunc = doc.Tables.Count;
            object listnumbers = doc.Lists.Count;
            object listpnumbers = doc.ListParagraphs.Count;
            object listtbumbers = doc.ListTemplates.Count;
            Lists lists = doc.Lists;
            ListParagraphs listps = doc.ListParagraphs;
            ListTemplates listts = doc.ListTemplates;
            object list1 = lists[1];
            object list2 = lists[2];
            object list3 = listps[1];
            object list4 = listts[1];
            string k = listps[3].Range.Text.Trim();
            //object level = lists[1].ApplyListTemplateWithLevel
            string[] k3 = new string[3];
            for (int i = 0; i <= 1; i++)
            {
                k3[i] = lists[i + 1].Range.Text.Trim();
            }
            int[] num3 = new int[2];
            for (int i = 0; i <= 1; i++)
            {
                num3[i] = lists[i + 1].Range.Start;
            }
            string[] k4 = new string[3];
            for (int i = 0; i <= 2; i++)
            {
                k3[i] = listps[i + 1].Range.Text.Trim();
            }
            listpnumbers = doc.ListParagraphs.Count;
            app.Documents.Close(ref unknow, ref unknow, ref unknow);
            app.Quit(ref unknow, ref unknow, ref unknow);
            return null;
        }


    }



}


