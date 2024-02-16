

// See https://aka.ms/new-console-template for more information
using System.Formats.Asn1;
using System.Globalization;
using System.Text.RegularExpressions;
using IronPdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Runtime.Serialization;
using static System.Net.Mime.MediaTypeNames;
using System.Data;
using System.Diagnostics;
using SixLabors.ImageSharp.Metadata.Profiles.Exif;
using System.Diagnostics.CodeAnalysis;

//Licence Iron PDF (test - 1 mois)
License.LicenseKey = "IRONSUITE.PONGLA.PUBLIC.LAPOSTE.NET.23963-6A3FDD8D07-JLJAD-OLDZZKIU5GM2-BMN4NAAKREBO-574WYQAE4LHC-XPWT7S72AQSZ-GEXT6IZAHB7S-ASPCER43TYTZ-ZHAJKU-TLUTDCDRVWKLUA-DEPLOYMENT.TRIAL-3XMQDE.TRIAL.EXPIRES.25.FEB.2024";

Console.WriteLine("C'est parti !");

var sourceFolderPath = "\\\\SAPIENCE12\\Partage\\Pierre Wattenne\\Documents justificatifs\\Relevés de compte"; //\\La Poste\\2020"; //\\Compte Carte CB;\\La Poste \\Société Générale

Merge(sourceFolderPath);

string rootPath;

void Merge(string sourceFolderPath)
{
    rootPath = sourceFolderPath;
    MergeLoc(sourceFolderPath, 0);
}

void MergeLoc(string sourceFolderPath, int level)
{
    #region Local functions

    void ConsoleWrite(string msg, int? forcedLevel = null) => Console.WriteLine($"{new string(' ', 2 * (forcedLevel ?? level))}{msg}");

    string GetMergedFileName(string currentFolderPath)
    {
        var localPath = currentFolderPath.Replace(rootPath, "");
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


