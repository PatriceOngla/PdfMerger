using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronPdf;
using System.Diagnostics;
using System.ComponentModel.Design;

namespace GetBankccountData
{


    internal class PdfMerger
    {
        public string RootPath { get; private set; }

        public void Merge(string sourceFolderPath)
        {
            RootPath = sourceFolderPath; 
            MergeLoc(sourceFolderPath, 0);
        }

        public void MergeLoc(string sourceFolderPath, int level)
        {
            #region Local functions

            void ConsoleWrite(string msg, int? forcedLevel = null) => Console.WriteLine($"{new string(' ', 2 * (forcedLevel ?? level))}{msg}");

            string GetMergedFileName(string currentFolderPath)
            {
                var localPath = currentFolderPath.Replace(RootPath, "");  
                if (localPath.StartsWith("\\")) localPath = localPath.Substring(1);
                var r = localPath.Replace(@"\", "-") + ".pdf";
                return r;
            }  

            string? MergeLocalPdfs(string currentFolderPath)
            {
                Console.WriteLine();
                ConsoleWrite($"Start merging PDF documents in subfloder './{Path.GetFileName(currentFolderPath)}'.", level + 1);
                string[] pdfNames = Directory.GetFiles(currentFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                if (pdfNames.Length == 0)
                {
                    ConsoleWrite($"No PDF documents found in '{currentFolderPath}'.", level + 1);
                    return null;
                }

                var pdfDocs = new List<PdfDocument>();
                foreach (var pdfName in pdfNames)
                {
                    ConsoleWrite($"Found pdf file '{Path.GetFileName(pdfName)}'", level + 2);
                    var pdf = PdfDocument.FromFile(pdfName);
                    pdfDocs.Add(pdf);
                }
                using var mergedDoc = PdfDocument.Merge(pdfDocs);
                var mergedDocName = GetMergedFileName(currentFolderPath);
                var mergedDocFullName = Path.Combine(currentFolderPath, mergedDocName);
                mergedDoc.SaveAs(mergedDocFullName);
                ConsoleWrite($"{pdfDocs.Count} documents merged in '{mergedDocName}'", level + 2);
                return mergedDocName;
            }

            #endregion

            string[] subFolders = Directory.GetDirectories(sourceFolderPath);
            Console.WriteLine();
            ConsoleWrite($"Start merging PDF files from '{sourceFolderPath}'.");

            if (subFolders.Length == 0)
                ConsoleWrite($"No subfolders.");
            else
            {
                foreach (var folderPath in subFolders)
                {
                    var mergedFileName = MergeLocalPdfs(folderPath);
                    ConsoleWrite($"Start merging files in subdirectories...", level + 1);
                    MergeLoc(folderPath, level + 1);
                }
            }
            ConsoleWrite($"PDF files merged in '{sourceFolderPath}'.");

        }




    }
}
