using MissionPlanner.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Community.CsharpSqlite.Sqlite3;
using Excel = Microsoft.Office.Interop.Excel;

namespace MissionPlanner
{
    internal class NavWrite
    {
        private bool IsEmpty(object cellValue)
        {
            if (cellValue == null)
                return true;

            // Vérifie si c'est un objet COM
            if (cellValue.GetType().IsCOMObject)


            {
                try
                {
                    // Si possible, essaye de récupérer la valeur sous-jacente de l'objet COM
                    dynamic comObject = cellValue;
                    var comValue = comObject.Value; // Cela dépend de la structure de l'objet COM
                    return comValue == null || string.IsNullOrWhiteSpace(comValue.ToString());
                }
                catch
                {
                    // Si cela échoue, retourne 'false' ou 'true' selon ce que tu veux par défaut
                    return true;  // On considère l'objet COM comme vide si la lecture échoue
                }
            }

            // Vérifie si c'est une chaîne vide ou composée d'espaces blancs
            return string.IsNullOrWhiteSpace(cellValue.ToString());
        }

        public void Write(int curr_wp)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Vérifier si Excel est ouvert
                excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                int i = curr_wp;

                // Vérifier si une seule feuille est ouverte
                if (excelApp.Workbooks.Count != 1)
                {
                    Console.WriteLine("Il doit y avoir exactement une feuille Excel ouverte.");
                    return;
                }

                // Obtenir la première feuille
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                System.DateTime dateTime = MainV2.comPort.MAV.cs.datetime;
                string heure = dateTime.ToString("HH:mm:ss");

                // Vérifier si les cellules sont vides avant d'écrire
                if (IsEmpty(worksheet.Cells[5 + i, "M"]) && IsEmpty(worksheet.Cells[5 + i, "Q"]) && IsEmpty(worksheet.Cells[5 + i, "U"]))
                {
                    // Écrire dans les cellules spécifiées
                    worksheet.Cells[5 + i, "M"] = heure;
                    worksheet.Cells[5 + i, "Q"] = MainV2.comPort.MAV.cs.battery_usedmah;
                    worksheet.Cells[5 + i, "U"] = MainV2.comPort.MAV.cs.battery_voltage;
                }
                else
                {
                    // Afficher le contenu des cellules dans la console pour déboguer
                    Console.WriteLine("Contenu de la cellule M: " + worksheet.Cells[5 + i, "M"]);
                    Console.WriteLine("Contenu de la cellule Q: " + worksheet.Cells[5 + i, "Q"]);
                    Console.WriteLine("Contenu de la cellule U: " + worksheet.Cells[5 + i, "U"]);
                }
                if (i == 1)
                {
                    worksheet.Cells[15, "D"] = heure;
                    worksheet.Cells[14, "D"] = MainV2.comPort.MAV.cs.battery_usedmah;
                    worksheet.Cells[5 + i, "U"] = MainV2.comPort.MAV.cs.battery_voltage;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Une erreur s'est produite : " + ex.Message);
            }
            finally
            {
                // Libérer les ressources COM
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        public void WriteNavMission(List<Locationwp> mission, string worksheetName = null)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Récupère l’instance d’Excel en cours (lève si Excel pas ouvert)
                excelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                if (excelApp == null || excelApp.Workbooks.Count == 0)
                    throw new InvalidOperationException("Aucun classeur Excel ouvert.");

                if (excelApp.Workbooks.Count > 1)
                    throw new InvalidOperationException("Ouvre un seul classeur Excel pour éviter les ambiguïtés.");

                workbook = excelApp.ActiveWorkbook;

                if (workbook == null)
                    throw new InvalidOperationException("Aucun classeur actif.");

                // Vérifs générales
                if (workbook.ReadOnly)
                    throw new InvalidOperationException("Le classeur est en lecture seule.");
                if (workbook.ProtectStructure)
                    throw new InvalidOperationException("La structure du classeur est protégée.");
                if (workbook.ProtectWindows)
                    throw new InvalidOperationException("Les fenêtres du classeur sont protégées.");

                // Récupère une vraie Worksheet (pas une ChartSheet)
                Excel.Worksheet candidate = null;
                if (!string.IsNullOrEmpty(worksheetName))
                {
                    try { candidate = (Excel.Worksheet)workbook.Worksheets[worksheetName]; } catch { /* ignore */ }
                }
                if (candidate == null)
                {
                    // ActiveSheet peut être un graphique → on cherche la première Worksheet visible
                    if (workbook.ActiveSheet is Excel.Worksheet ws)
                        candidate = ws;
                    else
                    {
                        foreach (Excel.Worksheet w in workbook.Worksheets)
                        {
                            if (w.Visible == Excel.XlSheetVisibility.xlSheetVisible) { candidate = w; break; }
                        }
                    }
                }
                if (candidate == null)
                    throw new InvalidOperationException("Impossible de trouver une feuille de calcul visible.");

                worksheet = candidate;

                // Si la feuille est « protégée côté contenu », on tente une déprotection sans mot de passe
                // (à adapter si tu as un mot de passe)
                bool reProtectAfter = false;
                if (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectionMode)
                {
                    try
                    {
                        worksheet.Unprotect(Type.Missing);
                        reProtectAfter = true; // on réactivera plus bas en UserInterfaceOnly si besoin
                    }
                    catch
                    {
                        throw new InvalidOperationException("La feuille est protégée et ne peut pas être déprotégée (mot de passe ?).");
                    }
                }

                // Détection Protected View (si le fichier vient d’Internet, etc.)
                // NB: via Interop standard ce n’est pas toujours exposé proprement ; on détecte surtout ReadOnly
                // L’utilisateur doit « Activer la modification » si bandeau jaune.

                // C#
                int startRow = 5;
                int startCol = 5; // Colonne E
                int rows = mission.Count;
                int cols = 12;

                // Vérification que toutes les cellules sont vides
                bool allEmpty = true;
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = worksheet.Cells[startRow + r, startCol + c];
                        if (!IsEmpty(cell))
                        {
                            allEmpty = false;
                            break;
                        }
                    }
                    if (!allEmpty) break;
                }

                if (allEmpty)
                {
                    // Prépare le bloc à écrire
                    object[,] data = new object[rows, cols];
                    for (int i = 0; i < mission.Count; i++)
                    {
                        var wp = mission[i];
                        data[i, 0] = i;
                        data[i, 1] = 0;          // Current WP
                        data[i, 2] = wp.frame;
                        data[i, 3] = wp.id;
                        data[i, 4] = wp.p1;
                        data[i, 5] = wp.p2;
                        data[i, 6] = wp.p3;
                        data[i, 7] = wp.p4;
                        data[i, 8] = wp.lat;
                        data[i, 9] = wp.lng;
                        data[i, 10] = wp.alt;
                        data[i, 11] = 1;          // Autocontinue
                    }

                    Excel.Range topLeft = (Excel.Range)worksheet.Cells[startRow, startCol];
                    Excel.Range bottomRight = (Excel.Range)worksheet.Cells[startRow + rows - 1, startCol + cols - 1];
                    Excel.Range writeRange = worksheet.Range[topLeft, bottomRight];

                    try { writeRange.Locked = false; } catch { }

                    writeRange.Value2 = data;
                }
                else
                {
                    Console.WriteLine("Certaines cellules de la plage cible ne sont pas vides, aucune écriture effectuée.");
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Une erreur s'est produite : " + ex.Message);
                // Astuce de debug : afficher ces états
                try
                {
                    Console.WriteLine($"Workbook.ReadOnly={workbook?.ReadOnly}");
                    Console.WriteLine($"Workbook.ProtectStructure={workbook?.ProtectStructure}");
                    Console.WriteLine($"Workbook.ProtectWindows={workbook?.ProtectWindows}");
                    if (worksheet != null)
                    {
                        Console.WriteLine($"Worksheet.Name={worksheet.Name}");
                        Console.WriteLine($"Worksheet.ProtectContents={worksheet.ProtectContents}");
                        Console.WriteLine($"Worksheet.ProtectionMode={worksheet.ProtectionMode}");
                    }
                }
                catch { }
            }
            finally
            {
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}