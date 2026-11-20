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
    sheet["C1"].Value = "Fälligkeitsdatum";
    sheet["D1"].Value = "Priorität";
    sheet["E1"].Value = "Erledigt";
}


string[] liste = { "Titel:", "Beschreibung:", "Fälligkeitsdatum:", "Priorität:", "Erledigt:"};

string eingabe;
string wahl;
string[] aufgabe = new string[liste.Length];

do
{

    Console.WriteLine("=========Hauptmenü=========\n1 = Aufgabe erfassen\n2 = Aufgaben anzeigen\n3 = Aufgabe löschen\n4 = beenden\n5 = Aufgabe als erledigt markieren\n_____________________________________");
    wahl = Console.ReadLine();

    if (wahl == "1")
    {
        Console.WriteLine("Du hast Aufgabe erfassen gewählt.");


        for (int i = 0; i < (liste.Length -1); i++)
        {
            Console.WriteLine(liste[i]);
            eingabe = Console.ReadLine();
            Console.WriteLine("___________________");

            bool richtig = false;

            if (i == 2)
            {
                while (richtig == false)
                {
                    richtig = DateOnly.TryParse(eingabe, out DateOnly datum);

                    if (richtig == false)
                    {
                        Console.WriteLine("Kein Datum, nochmal:");
                        eingabe = Console.ReadLine();
                    }
                }
            }
            else if (i == 3)
            {
                int prio = 0;

                while (richtig == false)
                {
                    richtig = int.TryParse(eingabe, out prio);

                    if (prio < 1)
                    {
                        richtig = false;
                    }

                    if (prio > 3)
                    {
                        richtig = false;
                    }

                    if (richtig == false)
                    {
                        Console.WriteLine("1 bis 3, nochmal:");
                        eingabe = Console.ReadLine();
                    }
                }
                eingabe = prio.ToString();

            }

            aufgabe[i] = eingabe;
        }

        aufgabe[4] = "nein";
        //wird în die CSV Datei eingespeichert
        if (freeLine <= sheet.Rows.Count())
        {
            freeLine = sheet.Rows.Count() + 1;
        }

        sheet["A" + freeLine].Value = aufgabe[0];
        sheet["B" + freeLine].Value = aufgabe[1];
        sheet["C" + freeLine].Value = aufgabe[2];
        sheet["D" + freeLine].Value = aufgabe[3];
        sheet["E" + freeLine].Value = aufgabe[4];

        freeLine++;

        Console.WriteLine("Aufgabe wurde gespeichert.");
    }
    else if (wahl == "2")
    {
        Console.WriteLine("Du hast Aufgaben anzeigen gewählt.");
        Console.WriteLine(sheet);
    }
    else if (wahl == "3")
    {
        Console.WriteLine("Du hast Aufgabe löschen gewählt.");
        Console.WriteLine("Welche Aufgabe möchtest du löschen? (1,2,3...)");
        string loeschen = Console.ReadLine();
        
        int loeschenInt = System.Convert.ToInt32 (loeschen);
        

        sheet["A" + (loeschenInt + 1)].ClearContents();
        sheet["B" + (loeschenInt + 1)].ClearContents();
        sheet["C" + (loeschenInt + 1)].ClearContents();
        sheet["D" + (loeschenInt + 1)].ClearContents();
        sheet["E" + (loeschenInt + 1)].ClearContents();
        freeLine -= 1;
    }
    else if (wahl == "4")
    {
        Console.WriteLine("Programm wird beendet.");
        
    }
    else if (wahl == "5")
    {
        Console.WriteLine("Welche Aufgabe möchtest du als erledigt markieren? (1,2,3...)");
        string erledigt = Console.ReadLine();
        int erledigtInt = System.Convert.ToInt32 (erledigt);

        sheet["E" + erledigtInt].Value = "ja";

    }
    else
    {
        Console.WriteLine("Ungültige Eingabe");
    }

} while (wahl != "4");


workBook.SaveAsCsv("ToDo_List.csv", ",");
