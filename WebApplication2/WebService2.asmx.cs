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
using System.Diagnostics;
using System.Management.Instrumentation;
using System.Management;

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
        /*
        class Dictionary : IEquatable<Dictionary>
        {
            public bool Equals(Dictionary other)
            {
                return this.tag == other.tag;
            }

         }
         * */
        _Application app = new Microsoft.Office.Interop.Word.Application();
        public _Document doc;
        public _Document doc1;
        public _Document doc2;
        public _Document doc3;
        public _Document doc4;
        public _Document doc5;
        public _Document doc6;
        public _Document doc7;
        public string in_column;
        //各自的map体，互相独立开来

        public static ManagementObjectCollection mn = (new ManagementClass("Win32_Processor")).GetInstances();

        public Dictionary<string, string>[] map = new Dictionary<string, string>[25];
        public Dictionary<string, object>[] h = new Dictionary<string, object>[25];
        public static int doc_handler = 8;//总句柄
        public static int slice = 16;//线程数目
        int key_line = 25;//最大rs记录的block块长
        _Document[] docs_list = new _Document[25];
        _Application[] apps_list = new Microsoft.Office.Interop.Word.Application[25];
        public Dictionary<string, string> map1;
        public Dictionary<string, string> map2;
        Dictionary<string, string> map3;
        Dictionary<string, string> map4;
        Dictionary<string, string> map5;
        Dictionary<string, string> map6;
        Dictionary<string, string> map7;
        Dictionary<string, string> map8;
        Dictionary<string, object> h1;
        Dictionary<string, object> h2;
        Dictionary<string, object> h3;
        Dictionary<string, object> h4;
        Dictionary<string, object> h5;
        Dictionary<string, object> h6;
        Dictionary<string, object> h7;
        Dictionary<string, object> h8;
        public int pcount;
        public int tc_count;
        public string[] columns;
        public string root = "D:\\files\\words\\";//文件保存的路径;
        //public string root = "E:\\files\\";



        public class finaljson
        {

            //估计有上锁机制,导致会吧inman
            public List<Dictionary<string, string>> finalstrings = new List<Dictionary<string, string>>();
            //  public  List<Dictionary<string, string>> [] finals = new List<Dictionary<string, string>> [25];
            public List<Dictionary<string, object>> final_tc = new List<Dictionary<string, object>>();
        }
        public finaljson aaa = new finaljson();

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

        public void thread1(Document doc, int start, int end)//,ref List<Dictionary<string, string>> finalstrings)
        {


            string temp = null;
            // int start = 1;
            //int end = pcount;

            if (start <= 0) start = 1;
            if (end >= pcount) end = pcount;
            Dictionary<string, string> map = null;
            //  finalstrings=new List<Dictionary<string,string>>();
            Debug.WriteLine("inside {0}=>{1},{2}", doc.GetHashCode(), start, end);
            while (start <= end)
            {



                temp = doc.Paragraphs[start].Range.Text.Trim();


                if (map == null && Regex.Matches(temp, @"(^\[TSP[^\]]*?\]$)", RegexOptions.IgnoreCase).Count > 0)
                {
                    map = new Dictionary<string, string>(); map.Add("tag", temp);
                }
                else if (map != null)
                {
                    if (Regex.Matches(temp, @"^#.*?=.*?").Count > 0)
                    {


                        //大小写的问题尚未解决呢

                        Match mark = Regex.Match(temp, @"^#([^=]*?)=(.*?)$", RegexOptions.IgnoreCase);
                        // aaa.finalstrings.Add(  mark.Groups[1].Value+mark.Groups[2].Value);
                        if (Regex.Matches(in_column, mark.Groups[1].ToString().Trim(), RegexOptions.IgnoreCase).Count > 0)
                        {

                            map.Add(mark.Groups[1].ToString().Trim(), mark.Groups[2].Value.Trim());

                        }

                    }
                    else if (Regex.Matches(temp, @"^\[End\]$", RegexOptions.IgnoreCase).Count > 0)
                    {
                        if (!aaa.finalstrings.Contains(map))
                        {
                            aaa.finalstrings.Add(map);
                            map = null;
                        }//销毁对象}
                    }
                    else
                    {
                        //description字段哦,如果用户没有输入
                        if (in_column.Contains("description"))
                        {
                            if (map.ContainsKey("description")) { map["description"] += temp; }//1/(11):(map1.Add("description",temp));
                            else { map.Add("description", temp); }
                        }
                    }

                }//else map1==null
                else { }
                start++;







            }//while



        }//thread1


        [WebMethod]
        public string downfile(string doc_url)
        {

            string LocalPath = root + doc_url.Substring(doc_url.LastIndexOf('/'));
            FileStream fs = null;
            Stream sIn = null;
            HttpWebResponse wr = null;
            Uri u = new Uri(doc_url);

            HttpWebRequest mRequest = (HttpWebRequest)WebRequest.Create(u);
            mRequest.Method = "GET";
            mRequest.ContentType = "application/x-www-form-urlencoded";
            wr = (HttpWebResponse)mRequest.GetResponse();
            sIn = wr.GetResponseStream();
            fs = new FileStream(LocalPath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            byte[] bytes = new byte[4096];
            int start = 0;
            int length;
            while ((length = sIn.Read(bytes, 0, 4096)) > 0)
            {
                fs.Write(bytes, 0, length);
                start += length;
            }



            if (sIn != null) sIn.Close();
            if (wr != null) wr.Close();
            if (fs != null) fs.Close();

            return LocalPath;

        }

        [WebMethod(Description = "readtc_void")]
        public void readtc(Document doc, int start_line, int end_line)
        {



            start_line = start_line <= 0 ? 1 : start_line;
            end_line = end_line >= pcount ? pcount : end_line;
            Debug.WriteLine("inside {0}=>{1},{2}", doc, start_line, end_line);
            Microsoft.Office.Interop.Word.Table nowTable;
            Dictionary<string, object> h = null;
            for (int tablePos = start_line; tablePos <= end_line; tablePos++)
            {
                nowTable = doc.Tables[tablePos];
                Regex tsp_mark = new Regex(@"^(\[TSP-.*?-\d*?[^\r\n]*).*?");
                if (!tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Success)
                {
                    // aaa.finalstrings.Add(nowTable.Cell(1, 2).Range.Text.Trim());
                    continue;

                }
                // aaa.finalstrings.Add(columns.ToString());
                h = new Dictionary<string, object>();
                h.Add("tag", tsp_mark.Match(nowTable.Cell(1, 2).Range.Text.Trim()).Groups[1].Value);
                for (int rowPos = 1; rowPos <= nowTable.Rows.Count; rowPos++)
                {  //还要把中文给去掉,以及纵向合并的单元格


                    string flag = Regex.Match(nowTable.Rows[rowPos].Cells[1].Range.Text.Trim().ToLower().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase).Groups[1].Value;//,@"\w",RegexOptions.IgnoreCase).Groups[0].Value.Trim();
                    string[] flags = flag.Split(new char[] { ' ' });
                    flag = string.Join(" ", flags);
                    //   aaa.finalstrings.Add(flag);
                    bool mark = false;
                    foreach (string a in columns)
                    {
                        if (a.ToLower().Trim() == flag) { mark = true; break; }

                    }

                    //aaa.finalstrings.Add(flag);

                    if (mark && nowTable.Rows[rowPos].Cells.Count == 2)
                    {

                        string text = nowTable.Rows[rowPos].Cells[2].Range.Text.Trim().Replace("\r", "").Replace("\u0007", "");
                        //  aaa.finalstrings.Add(flag + "haha here" + nowTable.Rows[rowPos].Cells[1].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                        if (rowPos != 1) { h.Add(flag, text); continue; }
                        //此时要解析出description,source,safety
                        h.Add("Source", null);
                        MatchCollection matches = Regex.Matches(text, @"\[Source:([^\]]*?\])\]", RegexOptions.IgnoreCase);
                        foreach (Match match in matches)
                        {
                            GroupCollection groups = match.Groups;
                            h["Source"] = h.ContainsKey("Source") ? (h["Source"] += (groups[1].ToString() + ",")) : null;

                            // h.Add((i++).ToString(), groups[1].Value);
                        }

                        if (h["Source"] != null) h["Source"] = h["Source"].ToString().Substring(0, h["Source"].ToString().Length - 1);
                        Match match_safety = Regex.Match(text, @"\[Safety:([^\]]*?)\]", RegexOptions.IgnoreCase);
                        if (match_safety.Success) h.Add("Safety", match_safety.Groups[1].Value);
                        Match match_desc = Regex.Match(text, @"\]([^\[\]]*)\[", RegexOptions.IgnoreCase);
                        if (match_desc.Success) { h.Add(flag, match_desc.Groups[1].Value); }


                    }
                    else if (mark && nowTable.Rows[rowPos].Cells.Count > 2)
                    {
                        //处理tc_test的情况了的
                        //收集列名字
                        //   aaa.finalstrings.Add(flag + "tc_step here" + in_column.Trim());
                        int num = nowTable.Rows[rowPos].Cells.Count;
                        string[] column_name = new string[num + 1];

                        for (int i = 1; i <= nowTable.Rows[rowPos].Cells.Count; i++)
                        {

                            Match temp = Regex.Match(nowTable.Rows[rowPos].Cells[i].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""), @"[\u4e00-\u9fa5\/]*(.*)?", RegexOptions.IgnoreCase);
                            column_name[i] = temp.Groups[1].Value.ToLower();// +nowTable.Rows[rowPos].Cells[i].Range.Text.Trim();

                        }
                        column_name[0] = nowTable.Rows[rowPos].Range.Text.ToString();

                        Console.WriteLine(nowTable.Rows[rowPos].Cells[num].Range.Text);
                        //  column_name[1]="num";
                        h["test steps"] = null;
                        //往下收集列即可
                        int k = 0;
                        for (int start = rowPos + 1; start < nowTable.Rows.Count; start++)
                        {

                            if (nowTable.Rows[start].Cells.Count != num)
                            {
                                rowPos = start - 1; break;
                            }
                            Dictionary<string, string> tc_step = new Dictionary<string, string>();
                            int j;
                            for (j = 1; j <= num; j++)
                            {

                                tc_step.Add(column_name[j], nowTable.Rows[start].Cells[j].Range.Text.Trim().Replace("\r", "").Replace("\u0007", ""));
                                //  Console.WriteLine(tc_step[column_name[j]]);

                            }

                            h["test steps"] += (new JavaScriptSerializer().Serialize(tc_step)) + ",";


                        }//处理tc_step
                        h["test steps"] = "[" + h["test steps"].ToString().Substring(0, h["test steps"].ToString().Length - 1) + "]";
                    }//else if 
                }//for row

                aaa.final_tc.Add(h);

            }//for tables


        }












        [WebMethod(Description = "readtc")]
        public string resolve()
        {

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            string param = System.Web.HttpUtility.UrlDecode(HttpContext.Current.Request.Url.Query.ToString().Substring(1));
            string[] param_s = param.Split('&'); Hashtable key_values = new System.Collections.Hashtable(); ;
            foreach (string item in param_s)
            {
                string[] key_value = item.Split('=');
                //  this.GetType().GetField(key_value[0]).GetValue(key_value[1]).ToString();
                key_values.Add(key_value[0], key_value[1]);
            }


            String LocalPath = null;
            string column = key_values.ContainsKey("column") ? key_values["column"].ToString() : "", type = key_values.ContainsKey("type") ? key_values["type"].ToString() : "", doc_url = key_values.ContainsKey("doc_url") ? key_values["doc_url"].ToString() : "";
            if (column.Equals("") || type.Equals("") || doc_url.Equals(""))
                throw new Exception("输入参数不合法");
            //  return column;  

            string message = null;
            try
            {
                String savePath = downfile(doc_url);

                //不需要此步判断了 if (!File.Exists(savePath)) throw new Exception("保存文件失败" ）; 

                in_column = column;
                //WORD  中数据都规整为一个空格隔开来
                columns = column.Split(',');


                /*
                    _Application app2= new Microsoft.Office.Interop.Word.Application();
                    _Application app3= new Microsoft.Office.Interop.Word.Application();
                    _Application app4= new Microsoft.Office.Interop.Word.Application();
                    _Application app5= new Microsoft.Office.Interop.Word.Application();
                    _Application app6= new Microsoft.Office.Interop.Word.Application();
                    _Application app7= new Microsoft.Office.Interop.Word.Application();
                
             */
                object fileName = savePath;

                object unknow = System.Type.Missing;

                //目前一个线程再跑
                if (type == "rs") { delet_tables(savePath); }

                //count the paragraphs

                if (type == "rs")
                {
                    //然后完成对文档的解析

                    List<System.Threading.Tasks.Task> TaskList = new List<System.Threading.Tasks.Task>();
                    // 开启线程池,线程分配算法
                    System.Threading.Tasks.Task t = null;
                    int k = (int)Math.Ceiling((Double)slice / doc_handler);
                    for (int i = 0; i < doc_handler; i++)
                    {
                        Microsoft.Office.Interop.Word.Application app_in = new Microsoft.Office.Interop.Word.Application();
                        var doc1 = app_in.Documents.Open(ref fileName, ref unknow, true, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                        pcount = doc1.Paragraphs.Count;
                        docs_list[i] = (doc1);
                        apps_list[i] = (app_in);

                        int block = (int)Math.Ceiling((Double)pcount / slice);
                        for (int j = i * k + 1; j <= (i + 1) * k && j <= slice; j++)
                        {

                            //Debug.WriteLine("传入{0},{1}", pcount * (i) / slice + 1 - key_line, pcount * (i + 1) / slice + key_line);
                            int start_in = block * (j - 1) + 1 - key_line <= 0 ? 1 : block * (j - 1) + 1 - key_line, end_in = block * (j) + key_line >= pcount ? pcount : block * (j) + key_line;
                            Debug.WriteLine("outside{0}=>{1},{2}", doc1.GetHashCode(), start_in, end_in);
                            t = new System.Threading.Tasks.Task(() => thread1(doc1, start_in, end_in));//, ref aaa.finals[i]));
                            t.Start();
                            TaskList.Add(t);
                        }
                    }



                    /*     
                         var t1 = new System.Threading.Tasks.Task(() => thread1(1,pcount/8+key_line, map[0]));
                         var t2 = new System.Threading.Tasks.Task(() => thread1(pcount / 8 + 1 - key_line, pcount * 2 / 8 + key_line,  map[1]));
                         var t3 = new System.Threading.Tasks.Task(() => thread1(pcount * 2 / 8 + 1 - key_line, pcount * 3 / 8 + key_line,  map[2]));
                         var t4 = new System.Threading.Tasks.Task(() => thread1(pcount * 3 / 8 + 1 - key_line, pcount * 4 / 8 + key_line,  map[3]));
                         var t5 = new System.Threading.Tasks.Task(() => thread1(pcount * 4 / 8 + 1 - key_line, pcount * 5 / 8 + key_line,  map[4]));
                         var t6 = new System.Threading.Tasks.Task(() => thread1(pcount * 5 / 8 + 1 - key_line, pcount * 6 / 8 + key_line,  map[5]));
                         var t7 = new System.Threading.Tasks.Task(() => thread1(pcount * 6 / 8 + 1 - key_line, pcount * 7 / 8 + key_line,  map[6]));
                         var t8 = new System.Threading.Tasks.Task(() => thread1(pcount * 7 / 8 + 1 - key_line, pcount * 8 / 8 + key_line,  map[7]));
                         t1.Start(); TaskList.Add(t1);
                         t2.Start(); TaskList.Add(t2);
                         t3.Start(); TaskList.Add(t3);
                         t4.Start(); TaskList.Add(t4);
                         t5.Start(); TaskList.Add(t5);
                         t6.Start(); TaskList.Add(t6);
                         t7.Start(); TaskList.Add(t7);
                         t8.Start(); TaskList.Add(t8);
                 
                 */

                    System.Threading.Tasks.Task.WaitAll(TaskList.ToArray());//t1, t2, t3, t4, t5, t6, t7, t8);
                    var json = new JavaScriptSerializer().Serialize(aaa.finalstrings.Where((x, i) => aaa.finalstrings.FindIndex(z => z["tag"] == x["tag"]) == i).ToList());
                    message = json;// "{\"success\":true,\"msg\":" + (new JavaScriptSerializer().Serialize(json)) + "}";

                }//rs
                else if (type == "tc")
                {

                    //tc并发并没有什么问题


                    List<System.Threading.Tasks.Task> TaskList = new List<System.Threading.Tasks.Task>();

                    int k = (int)Math.Ceiling((Double)slice / doc_handler);
                    Debug.WriteLine("buchang {0}", k);
                    for (int i = 0; i < doc_handler; i++)
                    {
                        Microsoft.Office.Interop.Word.Application app_in = new Microsoft.Office.Interop.Word.Application();
                        var doc1 = app_in.Documents.Open(ref fileName, ref unknow, true, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                        ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                        pcount = doc1.Tables.Count;
                        docs_list[i] = (doc1);
                        apps_list[i] = (app_in);
                        int block = (int)Math.Ceiling((Double)pcount / slice);

                        for (int j = i * k + 1; j <= (i + 1) * k && j <= slice; j++)
                        {
                            //Debug.WriteLine("传入{0},{1}", pcount * (i) / slice + 1 - key_line, pcount * (i + 1) / slice + key_line);

                            int start_in = block * (j - 1) + 1 <= 0 ? 1 : block * (j - 1) + 1, end_in = block * (j) >= pcount ? pcount : block * (j);
                            Debug.WriteLine("outside {0},{1}", start_in, end_in);
                            System.Threading.Tasks.Task t = new System.Threading.Tasks.Task(() => readtc(doc1, start_in, end_in));
                            t.Start();
                            TaskList.Add(t);
                        }
                    }




                    System.Threading.Tasks.Task.WaitAll(TaskList.ToArray());
                    var json = new JavaScriptSerializer().Serialize(aaa.final_tc);
                    //(new HashSet<Dictionary<string, object>>(aaa.final_tc)));
                    message = json;// "{\"success\":true,\"msg\":" + (new JavaScriptSerializer().Serialize(json)) + "}";



                }


            }

            catch (Exception e)
            {

                message = e.ToString();//
                // "{\"success\":false,\"msg\":\"" + e.Message + e.StackTrace + e.TargetSite + "\"}";

                return message;


            }

            finally
            {
                stopwatch.Stop();

                //  return  json;

                //异不异常到最后都关闭文档,避免word一直处于打开状态占用资源
                object unknows = System.Type.Missing;
                Debug.WriteLine("打开大小为{0}", docs_list.Length);
                //   if (doc != null) doc.Close();
                for (int i = 0; i < docs_list.Length; i++)
                {
                    if (docs_list[i] == null) { break; } Debug.WriteLine("开始close"); docs_list[i].Close(ref unknows, ref unknows, ref unknows); Debug.WriteLine("结束close");
                }
                for (int i = 0; i < apps_list.Length; i++)
                {

                    if (apps_list[i] == null) { break; } Debug.WriteLine("开始quit"); apps_list[i].Quit(ref unknows, ref unknows, ref unknows); Debug.WriteLine("结束quit");
                }

                GC.Collect();
                GC.Collect();
                Context.Response.ContentType = "text/json";
                // Context.Response.Write(stopwatch.Elapsed);
                Context.Response.Write(message);
                // Context.Response.Write(json);

                Context.Response.End();
                /*  doc1.Close(ref unknow, ref unknow, ref unknow);
                 foreach (_Application apps in apps_list)
                  doc2.Close(ref unknow, ref unknow, ref unknow);

                  doc3.Close(ref unknow, ref unknow, ref unknow);

                  doc4.Close(ref unknow, ref unknow, ref unknow);

                  doc5.Close(ref unknow, ref unknow, ref unknow);

                  doc6.Close(ref unknow, ref unknow, ref unknow);

                  doc7.Close(ref unknow, ref unknow, ref unknow);
    */

                /*     app1.Quit(ref unknow, ref unknow, ref unknow);
                     app2.Quit(ref unknow, ref unknow, ref unknow);
                     app3.Quit(ref unknow, ref unknow, ref unknow);
                     app4.Quit(ref unknow, ref unknow, ref unknow);
                     app5.Quit(ref unknow, ref unknow, ref unknow);
                     app6.Quit(ref unknow, ref unknow, ref unknow);
                     app7.Quit(ref unknow, ref unknow, ref unknow);
                 */







            }//finally
            return null;//不是正常请求方式,不可见这结果
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


