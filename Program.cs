﻿using System.Web;
using System.Diagnostics;
using HtmlAgilityPack;
using ClosedXML.Excel;

//Create settings files with base values
if (!File.Exists("login.txt"))
{
    File.WriteAllText("login.txt", "username=\npassword=");
    Console.WriteLine("Please fill in login details in the newly created 'login.txt' file and restart the program\nPress any key to exit..");
    Console.ReadKey(true);
    goto endofprogram;
}
if (!File.Exists("options.txt"))
{
    File.WriteAllText("options.txt", "user_agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0\nsave_grades=true\nverbose=true\nlist=true\nrecord_performance=false\nprogress_count=true\nsave_excel=true\nexcel_debug=false");
}

//Variables
string[] login = File.ReadAllLines("login.txt");
string username = login[0].Split('=')[1];
string password = login[1].Split('=')[1];
Console.WriteLine(username);
Console.WriteLine(password);
string auth = "";
string session_info = "";
string base_url = "https://aplikace.skolaonline.cz/SOL/";
string raw_html = "";
string[] excel_columns = new string[6] { "A", "B", "C", "D", "E", "F" };
int node_count;
int grade_count = 1;
int excel_column_count = 0;
int excel_row_count = 1;
bool auth_status = false;
bool already_failed = false;
Stopwatch sw = new();
IEnumerable<string> cookie;
HtmlDocument htmlDoc = new();
XLWorkbook xls = new();
IXLWorksheet sheet = xls.AddWorksheet("Sheet 1");
IXLCell excel_current_cell;

//Options - i know its ugly
string[] options = File.ReadAllLines("options.txt");
string user_agent = options[0].Split('=')[1];
bool save_grades = Convert.ToBoolean(options[1].Split('=')[1]);
bool verbose = Convert.ToBoolean(options[2].Split('=')[1]); ;
bool list = Convert.ToBoolean(options[3].Split('=')[1]); ;
bool record_performance = Convert.ToBoolean(options[4].Split('=')[1]); ;
bool progress_count = Convert.ToBoolean(options[5].Split('=')[1]); ;
bool save_excel = Convert.ToBoolean(options[6].Split('=')[1]); ;
bool excel_debug = Convert.ToBoolean(options[7].Split('=')[1]); ;

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//This function repeats 3 times in the code, which used to take up way more lines than needed
void column_parse(string column)
{

    if (save_grades)
    {
        File.AppendAllText("grades.txt", column);
    }
    if (verbose)
    {
        Console.Write(column);
    }
    if (save_excel)
    {
        excel_current_cell = sheet.Cell($"{excel_columns[excel_column_count % 6]}{excel_row_count}");
        if (excel_debug)
        {
            Console.Write($"{excel_columns[excel_column_count % 6]}{excel_row_count} ");
        }
        //Doubles and numbers are properly formatted
        excel_current_cell.Value = column;
        if (node_count == 7 || node_count == 8)
        {
            try
            {
                excel_current_cell.Value = Double.Parse(column.Replace(".", ","));
            }
            catch
            {
                excel_current_cell.Value = int.Parse(column.Replace(".", ","));
            }
            excel_current_cell.Style.NumberFormat.NumberFormatId = 0;
        }
        excel_column_count++;
    }
    node_count++;
}

while (!auth_status)
{
    using (var client = new HttpClient())
    {
        //Login POST
        using (var request = new HttpRequestMessage())
        {
            //Request Payload
            var values = new Dictionary<string, string>()
            {
                {"__EVENTTARGET","dnn$ctr994$SOLLogin$btnODeslat"},
                {"__EVENTARGUMENT",""},
                {"__VIEWSTATE",""},
                {"__VIEWSTATEGENERATOR",""},
                {"__VIEWSTATEENCRYPTED",""},
                {"__PREVIOUSPAGE",""},
                {"__EVENTVALIDATION",""},
                {"dnn$dnnSearch$txtSearch",""},
                {"JmenoUzivatele",$"{username}"},
                {"HesloUzivatele",$"{password}"},
                {"ScrollTop",""},
                {"__dnnVariable",""},
                {"__RequestVerificationToken",""},
            };

            Console.WriteLine("Attempting login...");

            //Request Properties
            request.Content = new FormUrlEncodedContent(values);
            request.Method = HttpMethod.Post;
            request.RequestUri = new Uri(base_url + "Prihlaseni.aspx");
            request.Headers.Add("User-Agent", user_agent);

            HttpResponseMessage response = await client.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();

            //Retrieval of '.ASPXAUTH' and other cookies, the first being a token for authorization         
            response.Headers.TryGetValues("set-cookie", out cookie);
            already_failed = false;
            try
            {
                auth = cookie.ElementAt(0);
                session_info = cookie.ElementAt(1);
            }
            catch
            {
                Console.WriteLine("Login failed. Try again.");
                auth_status = false;
                already_failed = true;
            }

            //If login is denied, only one 'Set-Cookie' header is returned, thus both elements of the 'cookie' variable end up being the same
            if (auth == session_info)
            {
                if (!already_failed) { Console.WriteLine("Login failed. Try again."); }
                auth_status = false;
            }
            else
            {
                Console.WriteLine("Login successful.");
                auth_status = true;
            }

            File.WriteAllText("response1.html", content);
        }
        //Grades GET    TODO: Still haven't accounted for these lists having more pages, as I can't test it myself.
        using (var request = new HttpRequestMessage())
        {
            //Request Properites
            request.Method = HttpMethod.Get;
            request.RequestUri = new Uri(base_url + "App/Hodnoceni/KZH003_PrubezneHodnoceni.aspx");
            request.Headers.Add("User-Agent", user_agent);
            request.Headers.Add("Cookie", $"{auth}");
            request.Headers.Add("Cookie", $"{session_info}");

            HttpResponseMessage response = await client.SendAsync(request);
            string content = await response.Content.ReadAsStringAsync();
            raw_html = content;

            File.WriteAllText("response2.html", content);
        }
    }
}

htmlDoc.LoadHtml(raw_html);
HtmlNodeCollection nodes = htmlDoc.DocumentNode.SelectNodes("//td/div/table/tbody/tr");

//Used for performance diagnostic
if (record_performance)
{
    sw.Start();
}

foreach (var HtmlNode in nodes)
{
    if (!list)
    {
        Console.Clear();
    }
    if (progress_count)
    {
        Console.Write($"{grade_count}/{nodes.Count} ");
    }
    node_count = 0;
    foreach (var node in HtmlNode.ChildNodes)
    {
        //In the HTML, the table has way more td and th elements which are invisible and contain nothing, this filters them out
        if (node_count < 4)
        {
            node_count++;
            continue;
        }
        else if (node_count > 9)
        {
            node_count++;
            continue;
        }
        //Columns in table either have a 'uv' value or 'title' value
        try
        {
            column_parse(HttpUtility.HtmlDecode(node.Attributes["uv"].Value) + " ");
        }

        //NullReferenceException is invoked basically every time
        catch (NullReferenceException)
        {
            try
            {
                //If another NullReferenceException is caught here, the column has neither a 'uv' or a 'title' and is an empty column
                column_parse(HttpUtility.HtmlDecode(node.Attributes["title"].Value + " "));
            }
            catch
            {
                column_parse("");
            }
        }
    }

    if (save_grades)
    {
        File.AppendAllText("grades.txt", "\n");
    }
    if (list)
    {
        Console.Write("\n");
    }

    excel_row_count++;
    grade_count++;
}

//Printing optional alerts
if (save_grades)
{
    Console.WriteLine("Exported to text file.");
}
if (save_excel)
{
    xls.SaveAs("grades.xlsx");
    Console.WriteLine("Exported to Excel File.");
}
if (record_performance)
{
    sw.Stop();
    Console.WriteLine("\nTotal runtime: " + sw.Elapsed.TotalMilliseconds + " milliseconds");
}

//Press key to exit
Console.Read();
endofprogram:;