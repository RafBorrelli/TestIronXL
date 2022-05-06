using IronXL;
using System.Data;
using System.Diagnostics;

Console.WriteLine("Hello, World!");
Console.WriteLine();

byte[] result;
DataSet ds = new();
DataTable t = new();
ds.Tables.Add(t);

List<String> columnHeaderList = new();

//Fill data table with fake data
for (int i = 0; i < 10; i++)
{
    //Here we give a name to the DataTable's columns.
    //Would be great if IronXl could automatically retrieve them and use them as the first line of the excel file.
    t.Columns.Add($"Col{i}", typeof(String));

    columnHeaderList.Add($"Col{i}");    //We fill a list with colum header names. This can be avoided if IronXl could grab the names from the DataTable's columns captions
}

for (int x = 0; x < 1000; x++)
{
    DataRow row = t.NewRow();
    for (int i = 0; i < t.Columns.Count; i++)
    {
        row[$"Col{i}"] = $"row {x} col {i}";    
    }
    t.Rows.Add(row);
}


License.LicenseKey = "IRONXL.TESTUTENTE.11748-326EEDD90F-BLGNAGAXFZOLDLN-77SSGAI7WCXE-5Y6TK5CZBK3V-SGEMTQ623TJ4-XF7YXV2TYCUR-SLUMZO-TWWQ6X5XP6KEEA-DEPLOYMENT.TRIAL-JF7FRZ.TRIAL.EXPIRES.09.MAR.2022";


Stopwatch timer = new();


//METHOD 1 - DYNAMIC / SLOW
Console.WriteLine("Start Method 1");
timer.Start();

WorkBook wb = WorkBook.Load(ds);
var ws = wb.DefaultWorkSheet;
ws.InsertRow(0);
for (int i = 0; i < columnHeaderList.Count; i++)
{
    ws.Rows[1].Columns[i].Value = columnHeaderList[i];
}

timer.Stop();
Console.WriteLine($"Method 1 Time elapsed: {timer.Elapsed}");

//METHOD 1 - DYNAMIC / SLOW
Console.WriteLine();
Console.WriteLine("Start Method 1 2nd time (give it a 2nd chance)");
timer.Restart();

WorkBook wb1 = WorkBook.Load(ds);
var ws1 = wb1.DefaultWorkSheet;
ws1.InsertRow(0);
for (int i = 0; i < columnHeaderList.Count; i++)
{
    ws1.Rows[1].Columns[i].Value = columnHeaderList[i];
}

timer.Stop();
Console.WriteLine($"Method 1 Time elapsed: {timer.Elapsed}");


//METHOD 2 - STATIC
Console.WriteLine();
Console.WriteLine("Start Method 2");
timer.Restart();

WorkBook wb2 = WorkBook.Load(ds);
var ws2 = wb2.DefaultWorkSheet;
ws2.InsertRow(0);

ws2["A1"].Value = columnHeaderList[0];
ws2["B1"].Value = columnHeaderList[1];
ws2["C1"].Value = columnHeaderList[2];
ws2["D1"].Value = columnHeaderList[3];
ws2["E1"].Value = columnHeaderList[4];
ws2["F1"].Value = columnHeaderList[5];
ws2["G1"].Value = columnHeaderList[6];
ws2["H1"].Value = columnHeaderList[7];
ws2["I1"].Value = columnHeaderList[8];
ws2["J1"].Value = columnHeaderList[9];

timer.Stop();
Console.WriteLine($"Method 2 Time elapsed: {timer.Elapsed}");
Console.WriteLine();
Console.WriteLine("Press enter to exit");
Console.Read();