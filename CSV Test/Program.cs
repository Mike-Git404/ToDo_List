using IronXL;
using System.Drawing.Printing;
using System.IO.Enumeration;
IronXL.License.LicenseKey = "IRONSUITE.M.KUONEN.IM26.STUD.BBBADEN.CH.5988-85CE31746B-B5IX7DMGBYDFKON3-HIKVEW4FEWPG-AAEJA2HTLAVI-STSZ6565T2CG-SCY5JSG7U3P5-DFQCUX2AOC4Q-JGLIWC-TEMFOV5YM36RUA-DEPLOYMENT.TRIAL-JOTWGQ.TRIAL.EXPIRES.18.DEC.2026\r\n";


string[] input = new string[4];
int freeLine = 2;
bool loop = true;

WorkBook workBook;
WorkSheet sheet;

//sucht den File Path zur genannten Datei
string fileName = "ToDo_List.Tasks.csv";
string filePath = Path.GetFullPath(fileName);

//Wenn die Datei mit dem Path wurde, dann wird keine neue Datei erstellt. Sonst schon
if (File.Exists(filePath) == true)
{
    //Problem: erstellt ein neues sheet, obwohl bereits eines vorhanden ist
    workBook = WorkBook.Load("ToDo_List.Tasks.csv");
    sheet = workBook.DefaultWorkSheet;
}
else
{
    workBook = WorkBook.Create();
    sheet = workBook.CreateWorkSheet("Tasks");

    sheet["A1"].Value = "Titel";
    sheet["B1"].Value = "Beschreibung";
    sheet["C1"].Value = "Priorität";
}



while (loop == true)
{
    Console.WriteLine("***Willst du?****");
    string start = Console.ReadLine();

    if (start == "ja")
    {
        Console.WriteLine("Titel?");
        input[0] = Console.ReadLine();

        Console.WriteLine("Beschreibung?");
        input[1] = Console.ReadLine();

        Console.WriteLine("Priorität?");
        input[2] = Console.ReadLine();

        if (freeLine <= sheet.Rows.Count())
        {
            freeLine = sheet.Rows.Count() + 1;
        }

        sheet["A" + freeLine].Value = input[0];
        sheet["B" + freeLine].Value = input[1];
        sheet["C" + freeLine].Value = input[2];

        freeLine++;

    }
    else if (start == "zeigen")
    {
        Console.WriteLine(sheet);
        

    }
    else if (start == "löschen")
    {
        sheet["A" + 1].ClearContents();
    }
    else
    {
        loop = false;
    }

}

workBook.SaveAsCsv("ToDo_List.csv", ",");