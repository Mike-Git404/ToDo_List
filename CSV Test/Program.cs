using IronXL;
using System.Data;
using System.Drawing.Printing;
using System.IO.Enumeration;
IronXL.License.LicenseKey = "IRONSUITE.M.KUONEN.IM26.STUD.BBBADEN.CH.5988-85CE31746B-B5IX7DMGBYDFKON3-HIKVEW4FEWPG-AAEJA2HTLAVI-STSZ6565T2CG-SCY5JSG7U3P5-DFQCUX2AOC4Q-JGLIWC-TEMFOV5YM36RUA-DEPLOYMENT.TRIAL-JOTWGQ.TRIAL.EXPIRES.18.DEC.2026\r\n";


string[] input = new string[4];


WorkBook workBook;
WorkSheet sheet;

//sucht den File Path zur genannten Datei
string fileName = "ToDo_List.Tasks.csv";
string filePath = Path.GetFullPath(fileName);

//Wenn die Datei mit dem Path wurde, dann wird keine neue Datei erstellt. Sonst schon
if (File.Exists(filePath) == true)
{
    
    workBook = WorkBook.Load("ToDo_List.Tasks.csv");
    sheet = workBook.DefaultWorkSheet;
}
else
{   // erstellt die Datei mit Spalten
    workBook = WorkBook.Create();
    sheet = workBook.CreateWorkSheet("Tasks");

    sheet["A1"].Value = "Titel";
    sheet["B1"].Value = "Beschreibung";
    sheet["C1"].Value = "Fälligkeitsdatum";
    sheet["D1"].Value = "Priorität";
    sheet["E1"].Value = "Erledigt";
}


string[] liste = { "Titel:", "Beschreibung:", "Fälligkeitsdatum:", "Priorität:", "Erledigt:"};
int freeLine = sheet.Rows.Count();

string eingabe;
string wahl;
string[] aufgabe = new string[liste.Length];

// läuft bis "4" eingegeben ist
do
{

    Console.WriteLine("=========Hauptmenü=========\n1 = Aufgabe erfassen\n2 = Aufgaben anzeigen\n3 = Aufgabe löschen\n4 = beenden\n5 = Aufgabe als erledigt markieren\n6 = Sortieren\n_____________________________________");
    wahl = Console.ReadLine();

    //Aufgabe erfassen
    if (wahl == "1")
    {
        Console.WriteLine("Du hast Aufgabe erfassen gewählt.");


        for (int i = 0; i < (liste.Length - 1); i++)
        {
            Console.WriteLine(liste[i]);
            eingabe = Console.ReadLine();
            Console.WriteLine("___________________");

            bool richtig = false;

            if (i == 2)
            {   //überprüft ob die Eingabe ein gültiges Datum ist
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
            // überprüft, ob die Priorität gültigt ist
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
    // Aufgaben Anzeigen
    {
        Console.WriteLine("Du hast Aufgaben anzeigen gewählt.");
        Console.WriteLine(sheet);
    }
    else if ((wahl == "3") && (freeLine - 1 >= 2))
    // Azfgabe löschen
    {
        Console.WriteLine("Du hast Aufgabe löschen gewählt.");
        Console.WriteLine("Welche Aufgabe möchtest du löschen? (1,2,3...)");
        string loeschen = Console.ReadLine();

        int loeschenInt = System.Convert.ToInt32(loeschen);

        sheet["A" + (loeschenInt + 1)].ClearContents();
        sheet["B" + (loeschenInt + 1)].ClearContents();
        sheet["C" + (loeschenInt + 1)].ClearContents();
        sheet["D" + (loeschenInt + 1)].ClearContents();
        sheet["E" + (loeschenInt + 1)].ClearContents();
        freeLine -= 1;

        Console.WriteLine("Aufgabe wurde gelöscht!");
    }
    else if (wahl == "4")
    //Programm beenden
    {
        Console.WriteLine("Programm wird beendet.");

    }
    else if (wahl == "5")
    // Aufgabe als erledigt Markieren
    {
        Console.WriteLine("Welche Aufgabe möchtest du als erledigt markieren? (1,2,3...)");
        string erledigt = Console.ReadLine();
        int erledigtInt = System.Convert.ToInt32(erledigt);

        sheet["E" + erledigtInt].Value = "ja";

    }
    else if (wahl == "6")
    {
        Console.WriteLine("Sortieren nach Datum (1) oder nach Priorität (2)sortieren?");
        string sort = Console.ReadLine();

        if (sort == "1")
        {
            var range = sheet[$"A2:E{freeLine}"];
            range.SortByColumn("C", SortOrder.Ascending);
            Console.WriteLine(sheet);
        }
        else if (sort == "2")
        {
            var range = sheet[$"A2:E{freeLine}"];
            range.SortByColumn("D", SortOrder.Ascending);
            Console.WriteLine(sheet);
        }
        else
        {
            Console.WriteLine("Ungültige Eingabe");
        }
    } 
    else
    {
        Console.WriteLine("Ungültige Eingabe");
    }

} while (wahl != "4");


workBook.SaveAsCsv("ToDo_List.csv", ",");