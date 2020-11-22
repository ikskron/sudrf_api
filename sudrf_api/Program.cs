using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using System.Net;
using System.Data.OleDb;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq.Expressions;
using Newtonsoft.Json.Linq;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Web;

namespace sudrf_api
{

    /*
    Для запроса лиц по делу используйте параметр caseParticipant, пример:

caseParticipant=Сидорова%20Валентина%20Ивановна

Также при указании этого параметра указывайте категорию дел с помощью параметра group.

Для СОЮ / Гражданские и административные дела - Гражданские и административные дела, пример:
group=Гражданские%20и%20административные%20дела    r_processType = 1

Для Мировые судьи / Гражданские и административные дела - Гражданские и административные дела мировых судей, пример:
group=Гражданские%20и%20административные%20дела%20мировых%20судей    r_processType = 2

Весь запрос будет выглядеть так:
https://parser-api.com/parser/sudrf_api/?key=КЛЮЧ&page=1&caseParticipant=Сидорова%20Валентина%20Ивановна&group=Гражданские%20и%20административные%20дела%20мировых%20судей



    */

    class caseid
    {
        public string id { get; set; }   ///// Идентификатор дела
        public string Person_id { get; set; }  // Ид персоны
       

    }
    class Program
    {
        static List<object> array = new List<object>(); //список дел, выводимых в таблицу
        static string dict = string.Empty; // справочник от сервиса
//       static int typ = -1;
       static string KEY = "4df1cd681456ed242c1cc2606371c7a4";
        static string connetionString = "Data Source=PROFIDATA;Initial Catalog=db_mess_bankruptcy;Integrated Security=SSPI;";
        static SqlConnection cnn = new SqlConnection(connetionString);
       static JObject js = null;
      // static ushort r_processType;
        static void Main(string[] args)
        {
            /*
            if (args.Length == 0)
            {
                
                return;
            }
          */


           // Console.WriteLine(args[0]);
            // Console.WriteLine(DateTime.Parse("23.03.2020 11:40", System.Globalization.CultureInfo.CreateSpecificCulture("ru-ru")).ToString("o", System.Globalization.CultureInfo.InvariantCulture));
         
            //  r_processType = Convert.ToUInt16(args[0]);


            //считывание справочника web
            /*
            JObject jt = JObject.Parse(@"{""getValuesForFields"":1}");
                       
            dict = req(KEY, jt);
            */



            // DataTable dt = GetData("select code, dsc from Filbert_court_dict where parent_name ='courtType'  order by parent_name, code");

            // Альтернатива - считывание справочника из БД

            string[] dictItems = { "courtType", "instance", "processType", "region" };
            js = JObject.Parse("{\"fields\":{}}");
            foreach (string item in dictItems)
            {
                DataTable dt = GetData("select dsc from Filbert_court_dict where parent_name = '" +item+"'  order by parent_name, code");
                dict = DataTableToJSONWithStringBuilder(dt, item);



                
                // JToken jt = (JToken)JObject.Parse(dict);


                ((JObject)js["fields"]).Add(new JProperty(item, Newtonsoft.Json.JsonConvert.DeserializeObject(dict)));

                //js.Add((dict));

            }
            //Console.WriteLine( get_dict("region",2));
            
            
                    detail_info();
            

        }



      static  void detail_info()
        {

            // Считываем входные параметры из таблицы
            string sql = "select top 10 id, name, caseNumber, person_id, r_region, cast(caseDateFrom as date) caseDateFrom, cast(caseDateTo as date) caseDateTo, caseNumber, hasDocument, r_processType from Filbert_court_requests where isnull(execution_flag,0)!=1 " +
             "order by id desc"   ;
            DataTable dt1 = GetData(sql);

            if (dt1.Rows.Count == 0) return;


            foreach (DataRow dt in dt1.Rows)            //Устарело, т.к. формируем запрос по последней записи
            {
                string jstr = @"{ ";
                jstr += (dt["caseNumber"] != DBNull.Value ? @"""caseNumber"": """ + dt["caseNumber"].ToString() + @"""," : string.Empty);                     //           @"""}" ; 


                //   jstr += (dt.Rows[0]["caseNumber"] != DBNull.Value ? @"""caseDateTo"": """ + dt.Rows[0]["caseNumber"].ToString() + @"""," : string.Empty);   //caseNumber
                jstr += (dt["hasDocument"] != DBNull.Value ? @"""hasDocument"":" + Convert.ToInt16(dt["hasDocument"]).ToString() + @"," : string.Empty);   //hasDocument

                // Для запросов по СОЮ и Мровым судьям, берутся следующие поля: Name (ФИО)


                if (dt["name"] != DBNull.Value)
                {
                    // Если задействованы эти запросы, остальные поля игнорируем
                    string fio = Regex.Replace(dt["Name"].ToString(), " ", "%26%23032%3B");  //   &#032;


                    jstr = @"{ ";
                    jstr += (dt["Name"] != DBNull.Value ? @"""caseParticipant"":""" + fio + @"""," : string.Empty);
                    jstr += Convert.ToInt16(dt["r_processType"].ToString()) == 1 ? @"""group"":""Гражданские%20и%20административные%20дела"","
                                                : @"""group"":""Гражданские%20и%20административные%20дела%20мировых%20судей""";

                }
                //суд, номер дела, инстанция, уровень суда, регион, статья или категория, судья, дата регистрации дела, дата окончания (с по), дата вступления решения в силу (с по), ФИО
                jstr += (dt["r_region"] != DBNull.Value ? @",""region"": """ + get_dict("region", Convert.ToInt32(dt["r_region"].ToString())) + @"""," : string.Empty);
                jstr += (dt["caseDateFrom"] != DBNull.Value ? @"""caseDateFrom"": """ + DateTime.Parse(dt["caseDateFrom"].ToString()).ToShortDateString() + @"""," : string.Empty);           //                caseDateFrom
                jstr += (dt["caseDateTo"] != DBNull.Value ? @"""caseDateTo"": """ + DateTime.Parse(dt["caseDateTo"].ToString()).ToShortDateString() + @"""," : string.Empty);   //caseDateTo 
                jstr += (dt["caseLegalForceDateFrom"] != DBNull.Value ? @"""caseLegalForceDateFrom"": """ + DateTime.Parse(dt["caseLegalForceDateFrom"].ToString()).ToShortDateString() + @"""," : string.Empty);   //caseLegalForceDateFrom	 
                jstr += (dt["caseLegalForceDateTo"] != DBNull.Value ? @"""caseLegalForceDateTo"": """ + DateTime.Parse(dt["caseLegalForceDateTo"].ToString()).ToShortDateString() + @"""," : string.Empty);   
                jstr += (dt["caseNumber"] != DBNull.Value ? @"""caseNumber"": """ + DateTime.Parse(dt["caseNumber"].ToString()).ToShortDateString() + @"""," : string.Empty);   

                











                //////////////////////////////////////////////////////////////////////////////


                // jstr = Regex.Replace(jstr, @"[\x2C](?=\\n$)", "");   
                jstr = Regex.Replace(jstr, @"[\x2C]($)", "");      //Убираем запятую последнюю ?
                jstr += @"}";


                // Постраничный перебор

                string j1 = req(KEY, JObject.Parse(jstr)); //первый запрос

                //////////////////////////////
                if (j1 == string.Empty || j1.Contains("error"))
                {
                    continue;
                    /*
                     Filbert_court_requests_log(2, Convert.ToInt16(dt["id"].ToString()), 0);
                     Console.WriteLine("\nКонец работы. Нажмите любую клавишу.");
                     Console.ReadKey();
                     return;
                    */
                }

                else if (Convert.ToInt32(JObject.Parse(j1)["total_count"].ToString()) == 0)
                {
                    string sql_i = "update Filbert_court_requests set execution_flag = 1 where id = " + dt["id"].ToString();
                    SqlCommand cmd_l = new SqlCommand(sql_i, cnn);
                    if (cnn.State != ConnectionState.Open) cnn.Open();
                    cmd_l.ExecuteNonQuery();
                    Filbert_court_requests_log(2, Convert.ToInt16(dt["id"].ToString()), 0);
                    continue;
                }
                else

                {
                    Filbert_court_requests_log(1, Convert.ToInt16(dt["id"].ToString()), 0);
                }
                ///////////////////////////
                ///
                add_caseid(j1, dt); // записали первый запрос
                if (Convert.ToInt32(JObject.Parse(j1)["total_count"].ToString()) > 200)
                {
                    int total = (Convert.ToInt32(JObject.Parse(j1)["total_count"].ToString()) - 200) / 10;

                    for (int i = 21; i <= total; i++) // Если первым запросом не выбрали, забираем остальные страницы (на странице по умолчанию 10 записей)
                    {
                        JObject jo1 = JObject.Parse(jstr);
                        ((JObject)jo1).Add(new JProperty("page", Newtonsoft.Json.JsonConvert.DeserializeObject(i.ToString())));
                        j1 = req(KEY, jo1);
                        add_caseid(j1, dt);
                    }

                }

                /// Далее запрос по каждому элементу с ид дела 
                /// 
                DataTable dt_p = GetData("select max(participants) m from Filbert_court_result_participants");
                int r_participants = Convert.ToInt32(dt_p.Rows[0]["m"] == DBNull.Value ? "0" : dt_p.Rows[0]["m"].ToString()) + 1;
                dt_p = GetData("select max(history) m from Filbert_court_result_history");
                int r_history = Convert.ToInt32(dt_p.Rows[0]["m"] == DBNull.Value ? "0" : dt_p.Rows[0]["m"].ToString()) + 1;
                dt_p = GetData("select max(documents) m from Filbert_court_result_documents");
                int r_doc = Convert.ToInt32(dt_p.Rows[0]["m"] == DBNull.Value ? "0" : dt_p.Rows[0]["m"].ToString()) + 1;

                foreach (caseid c in array)
                {
                    jstr = @"{""caseID"": """ + c.id + @"""}";
                    JObject jSon = JObject.Parse(jstr);
                    jstr = req(KEY, jSon);
                    if (jstr == string.Empty || jstr.Contains("error"))
                    {
                        Filbert_court_requests_log(2, Convert.ToInt16(dt["id"].ToString()), 1, c.id);
                        continue;
                    }
                    else
                    {

                        Filbert_court_requests_log(1, Convert.ToInt16(dt["id"].ToString()), 1, c.id);
                    }
                    jSon = JObject.Parse(jstr);
                    if (jSon["case"].HasValues == false) continue;
                    // запись в БД полная инфа по ИД
                    string sql_ins = "insert into Filbert_court_result (person_id, EndDate,RegisterDate, Category, r_Region, CourtName, processType, Judge, CaseNumber,Result, r_participants, " +
                        "r_history, r_documents,r_id) values (" +

                      (dt["person_id"] != DBNull.Value ? dt["person_id"].ToString() : "null") + "," +
                      (jSon["case"]["EndDate"] != null ? "'" + jSon["case"]["EndDate"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["RegisterDate"] != null ? "'" + jSon["case"]["RegisterDate"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["Category"] != null ? "'" + jSon["case"]["Category"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["Region"] != null ? getdict_id("region", jSon["case"]["Region"].ToString()) : "null") + "," +
                      (jSon["case"]["CourtName"] != null ? "'" + jSon["case"]["CourtName"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["ProcessType"] != null ? "'" + jSon["case"]["ProcessType"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["Judge"] != null ? "'" + jSon["case"]["Judge"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["CaseNumber"] != null ? "'" + jSon["case"]["CaseNumber"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["Result"] != null ? "'" + jSon["case"]["Result"].ToString() + "'" : "null") + "," +
                      (jSon["case"]["participants"] != null ? r_participants.ToString() : "null") + "," +
                      (jSon["case"]["history"] != null ? r_history.ToString() : "null") + "," +
                      (jSon["case"]["documents"] != null ? r_doc.ToString() : "null") +

                      ", " + dt["id"].ToString() +


                        ") ";
                    SqlCommand cmd_log = new SqlCommand(sql_ins, cnn);
                    if (cnn.State != ConnectionState.Open) cnn.Open();
                    cmd_log.ExecuteNonQuery();



                    if (jSon["case"]["history"] != null)
                    {
                        foreach (JObject j in jSon["case"]["history"])
                        {
                            // далее запись в связанные таблицы и запись лога
                            /////// Filbert_court_result_history
                            sql_ins = "insert into Filbert_court_result_history (history, Date, Name, result) values (" +
                             (r_history.ToString()) + "," +
                             (!string.IsNullOrEmpty(j["Date"].ToString()) ? "'" + (Regex.Replace(DateTime.Parse(j["Date"].ToString() + " " + j["Time"].ToString(), System.Globalization.CultureInfo.CreateSpecificCulture("ru-ru")).ToString("o", System.Globalization.CultureInfo.InvariantCulture), @"[\x2E](\d*$)", "")) + "'" : "null") + "," +
                             (j["Name"] != null ? "'" + j["Name"].ToString() + "'" : "null") + "," +
                             (j["Result"] != null ? "'" + j["Result"].ToString() + "'" : "null") +

                               ") ";

                            cmd_log = new SqlCommand(sql_ins, cnn);
                            if (cnn.State != ConnectionState.Open) cnn.Open();
                            cmd_log.ExecuteNonQuery();
                        }
                    }
                    /////////////////////////////////////////////////////////////
                    //////  Filbert_court_result_participants
                    if (jSon["case"]["participants"] != null)
                    {
                        foreach (JObject j in jSon["case"]["participants"])
                        {
                            // далее запись в связанные таблицы и запись лога
                            /////// Filbert_court_result_history
                            sql_ins = "insert into Filbert_court_result_participants (participants, type, Name, Category) values (" +
                             (r_participants.ToString()) + "," +
                             (j["Type"] != null ? "'" + j["Type"].ToString() + "'" : "null") + "," +
                             (j["Name"] != null ? "'" + j["Name"].ToString() + "'" : "null") + "," +
                             (j["Category"] != null ? "'" + j["Category"].ToString() + "'" : "null") +

                               ") ";

                            cmd_log = new SqlCommand(sql_ins, cnn);
                            if (cnn.State != ConnectionState.Open) cnn.Open();
                            cmd_log.ExecuteNonQuery();
                        }
                    }
                    ///////////////////////////////////////////////////////////////////
                    ///  Filbert_court_result_documents
                    ///  
                    if (jSon["case"]["documents"] != null)
                    {
                        foreach (JObject j in jSon["case"]["documents"])
                        {
                            sql_ins = "insert into Filbert_court_result_documents (documents, PublishDate, Type, Date, text, html) values (" +
                                (r_doc.ToString()) + "," +
                                (j["PublishDate"] != null ? "'" + Regex.Replace(DateTime.Parse(j["PublishDate"].ToString(), System.Globalization.CultureInfo.CreateSpecificCulture("ru-ru")).ToString("o", System.Globalization.CultureInfo.InvariantCulture), @"[\x2E](\d*$)", "") + "'" : "null") + "," +
                                (j["Type"] != null ? "'" + j["Type"].ToString() + "'" : "null") + "," +
                                (j["Date"] != null ? "'" + Regex.Replace(DateTime.Parse(j["Date"].ToString(), System.Globalization.CultureInfo.CreateSpecificCulture("ru-ru")).ToString("o", System.Globalization.CultureInfo.InvariantCulture), @"[\x2E](\d*$)", "") + "'" : "null") + "," +
                                                            (j["Text"] != null ? "'" + j["Text"].ToString() + "'" : "null") + "," +
                                                                   (j["Html"] != null ? "'" + j["Html"].ToString() + "'" : "null") +
                                  ") ";

                            cmd_log = new SqlCommand(sql_ins, cnn);
                            if (cnn.State != ConnectionState.Open) cnn.Open();
                            cmd_log.ExecuteNonQuery();

                        }
                    }

                   



                    r_participants++;
                    r_doc++;
                    r_history++;




                }

                string sqli = "update Filbert_court_requests set execution_flag = 1 where id = " + dt["id"].ToString();
                SqlCommand cmdl = new SqlCommand(sqli, cnn);
                if (cnn.State != ConnectionState.Open) cnn.Open();
                cmdl.ExecuteNonQuery();

            }

           
            



        }


        static string get_dict(string tok, int idx)
        {
            string str = string.Empty;

            JObject json = js;

           // if (json["fields"][tok][idx][idx.ToString()] ==null)             str = json["fields"][tok][idx].ToString();
            str = json["fields"][tok][idx].ToString().Replace(" ", "%20");

            


            return str;
        }

        static string getdict_id(string cat, string v)
        {

            for (int i1 = 0; i1 < js["fields"][cat].Count(); i1++)
            {
                if (js["fields"][cat][i1].ToString()==v)
                {
                    return i1.ToString();
                }

            }

            int i = js["fields"][cat].Count()+1;
            int parent = 0;

            switch (cat)
            {
                case "region":
                    parent = 1;
                    break;
                case "courtType":
                    parent = 2;
                    break;


            }




            string sql_ins = "insert into Filbert_court_dict (	code,	dsc,	parent_id,	parent_name) values (" + i.ToString() + 
            ", '" +  v  + "',"+ parent+", '"+cat+"')";
            SqlCommand cmd_log = new SqlCommand(sql_ins, cnn);
            if (cnn.State != ConnectionState.Open) cnn.Open();
            cmd_log.ExecuteNonQuery();




            return i.ToString() ;

        }


        private static DataTable GetData(string sqlCommand)
        {



            SqlCommand command = new SqlCommand(sqlCommand, cnn);
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = command;

            DataTable table = new DataTable();
            table.Locale = System.Globalization.CultureInfo.InvariantCulture;
            adapter.Fill(table);
            //            cnn.Close();
            return table;
        }



        private static string req(string key, JObject json = null)
        {

            //Используем полученные куки для запуска следующего метода

            // byte[] body = Encoding.Default.GetBytes(json.ToString());
            string par = "https://parser-api.com/parser/sudrf_api/?key=" + key + "&";

            // List<string> keys = new List<string>();
            // Формируем параметры строки


            var result = json.Descendants().OfType<JProperty>().Where(f => f is JProperty);
            //  keys.AddRange(result);
            foreach (JToken jt in result)
            {


                
                    par = par + jt.Path.ToString() + "=" + jt.ToObject<JProperty>().Value.ToString();
                
                // Console.Out.Write(jt);


                par += (result.Last<JToken>() == jt ? string.Empty : "&");



                /*
                par = par + jt.Path.ToString() + "=" + jt.ToObject<JProperty>().Value.ToString();

                par += (result.Last<JToken>() == jt ? string.Empty : "&");
                */

            }










            // Пример http://parser-api.com/parser/sudrf_api/?key=4df1cd681456ed242c1cc2606371c7a4&getValuesForFields=1


            // Cookie cookie = new Cookie(cookieVal[0].Split(new char[] { '=' })[0].ToString(), cookieVal[0].Split(new char[] { '=' })[1].ToString(), "/", "api.casebook.ru");

          //  string u = HttpUtility.UrlEncode(par);

        var httpWebRequest = (HttpWebRequest)WebRequest.Create(par);

            httpWebRequest.ContentType = @"application/json";
            httpWebRequest.Method = "GET";
            httpWebRequest.Host = "parser-api.com";
            httpWebRequest.Accept = @"*/*";
            httpWebRequest.Timeout = 240000;

            WebHeaderCollection myWebHeaderCollection = httpWebRequest.Headers;
          //  myWebHeaderCollection.Add("apikey: " + cookieVal[0]);

            //  httpWebRequest.CookieContainer = new CookieContainer();
            //  httpWebRequest.CookieContainer.Add(cookie);
            //  httpWebRequest.ContentLength = body.Length;
            /*
              using (Stream stream = httpWebRequest.GetRequestStream())
              {
                  stream.Write(body, 0, body.Length);
                  stream.Close();
              }

              */

            string orig = string.Empty;
            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                

                using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream(), Encoding.GetEncoding("UTF-8"), true))
                {
                    //ответ от сервера

                    orig = streamReader.ReadToEnd();
                  //  if (typ == 1) goto end;
                    ///////////////////////////////////////////ПИШЕМ ответ в файл//////////////////////////////////////////
                    System.IO.StreamWriter sw = null;
                    string path = DateTime.Now.ToString();
                    char[] denied = new[] { ':', '.', ' ' };
                    StringBuilder newString = new StringBuilder();
                    foreach (var ch in path)
                        if (!denied.Contains(ch))
                            newString.Append(ch);
                    path = newString.ToString();



                    path = "c:\\temp\\sudrf\\work\\sud_" + path + ".json";
                    System.IO.FileStream fs = new System.IO.FileStream(path, System.IO.FileMode.Append, System.IO.FileAccess.Write);     //, System.IO.FileShare.ReadWrite

                    sw = new System.IO.StreamWriter(fs, Encoding.GetEncoding(1251));
                    sw.WriteLine(orig);
                    sw.Close();

                    fs.Dispose();
                    fs.Close();

                    ///////////////////////////////////////////
                }
            }

            catch (Exception ex)
            {
                Console.Out.Write(ex.Message);

                if (ex.Message.Contains("(404)"))
                    //Console.Out.Write(ex.Message);
                    return "(404)";
                else if (ex.Message.Contains("(429)"))
                    return "(429)";
                else if (ex.Message.Contains("(400)"))
                {
                   
                    return "(400)";
                }

            }



           // end:


            return orig;





        }


        


        static public string DataTableToJSONWithStringBuilder(DataTable table, string elem)
     {
            //    var JSONString = new StringBuilder("{\"fields\": { \"" + elem + "\":");   
                var JSONString = new StringBuilder();   
            if (table.Rows.Count > 0)    
    {   
                

        JSONString.Append("[");   
        for (int i = 0; i<table.Rows.Count; i++)    
        {   
          //  JSONString.Append("{");   
            for (int j = 0; j<table.Columns.Count; j++)    
            {   
                if (j<table.Columns.Count - 1)    
                {
                            //  JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\",");   
                            JSONString.Append("\"" + table.Rows[i][j].ToString() + "\":");
                        }    
                else if (j == table.Columns.Count - 1)    
                {
                            // JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\"");   
                            JSONString.Append("\"" + table.Rows[i][j].ToString() + "\"");
                        }
            }   
            if (i == table.Rows.Count - 1)
                {
    // JSONString.Append("}");
                }
           else
                {
                        // JSONString.Append("},");
                        JSONString.Append(",");
                    }
        }   
JSONString.Append("]");
                // JSONString.Append('}',2);
//                JSONString.Append('}');
            }   
    return JSONString.ToString();

    }   



   static  void add_caseid(string j, DataRow dr )
        {
            if (JObject.Parse(j)["cases"] != null)
                {

                foreach (JObject ci in JObject.Parse(j)["cases"])
                {
                    caseid CaseId = new caseid();
                    CaseId.id = ci["Id"].ToString();
                    CaseId.Person_id = dr["person_id"].ToString();


                    array.Add(CaseId);


                }

            }
        }

static void Filbert_court_requests_log(int r_result, int request_id,  int typ_request, string caseid = "null")  
        {
            // Реализация лога

            // [request_id] --для typ_request=0 пишем: Filbert_court_request.id, для typ_request=1 пишем: Filbert_court_result.id
            // typ_request 0 - общий, 1- по caseid
            // casid для типа 1  -все идентификаторы дел
            /*
             // [r_result] 
            1	NULL	Результат получен
               2	NULL	Ошибка на стороне сервера
                 3	NULL	Ограничение, суточный лимит
                    4	NULL	Ограничение, месячный лимит 
             */


            string sql_i = "insert into Filbert_court_requests_log (r_result, request_id,  typ_request, caseid) values (" + r_result.ToString() + "," + request_id.ToString() +
                "," + typ_request.ToString() + ",'" + caseid + "')";

            SqlCommand c_log = new SqlCommand(sql_i, cnn);
            if (cnn.State != ConnectionState.Open) cnn.Open();
            c_log.ExecuteNonQuery();




        }



    }
}
