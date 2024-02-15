

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

//Licence Iron PDF (test - 1 mois)
License.LicenseKey = "IRONSUITE.PONGLA.PUBLIC.LAPOSTE.NET.23963-6A3FDD8D07-JLJAD-OLDZZKIU5GM2-BMN4NAAKREBO-574WYQAE4LHC-XPWT7S72AQSZ-GEXT6IZAHB7S-ASPCER43TYTZ-ZHAJKU-TLUTDCDRVWKLUA-DEPLOYMENT.TRIAL-3XMQDE.TRIAL.EXPIRES.25.FEB.2024";

Console.WriteLine("C'est parti !");

//Testé pour La Poste et SG mais pas pour Bred
var sourceFolderPath = "\\\\SAPIENCE12\\Partage\\Pierre Wattenne\\Documents justificatifs\\Relevés de compte\\La Poste\\2020"; //\\Compte Carte CB;\\La Poste \\Société Générale
var pdfs = Directory.GetFiles(sourceFolderPath, "*.pdf", SearchOption.AllDirectories);
var accountEntries = new List<AccountEntry>();

String accountName = "?";
String entryType = "?";

var datePattern = "";
var shortDatePattern = @"\d{2}\/(?'EntryMonth'\d{2})";
var longDatePattern = shortDatePattern + @"\/\d{2}";
var veryLongDatePattern = shortDatePattern + @"\/\d{4}";
var amountPattern = @"(?'Amount'(\d{1,3})(?:\.\d{3})*,\d{2})";
var spaceOrEndOfLinePattern = @"(?:\s|\r?\n|\r)";
var isCardAccount = false;

foreach (var pdf in pdfs)
{
    Console.WriteLine(pdf);

    DateTime fileEditionDate = GetFileEditionDate(Path.GetFileName(pdf));

    var pdfDocument = new IronPdf.PdfDocument(pdf);
    var text = pdfDocument.ExtractAllText();
    if (text == null) throw new Exception($"Unable to parse '{pdf}'.");

    var forSG = text.Contains("Société Générale");
    if (forSG)
    {
        accountName = "SG Patrice Ongla";
        isCardAccount = text.Contains("RELEVÉ CARTE");
        if (isCardAccount) entryType = "CB"; else entryType = "Autre";
    }
    else
    {
        accountName = "BP Claire Pellerin";
        entryType = "?";
    }

    SetSubPatterns(forSG, entryType);

    var textToParse = GetTextToParse(text, forSG, entryType);

    var pattern = GetPattern(forSG, entryType);
    Regex regex = new Regex(pattern, RegexOptions.Singleline | RegexOptions.Multiline);

    var matches = regex.Matches(textToParse);

    if (matches.Count == 0)
        throw new Exception("No match !");

    foreach (Match match in matches)
    {
        var rowEffectiveDate = match.Groups["EffectiveDate"].Value;
        var rowValueDate = match.Groups["ValueDate"].Value;
        string effectiveDate, valueDate;
        if (forSG)
        {
            effectiveDate = rowEffectiveDate;
            valueDate = rowValueDate;
        }
        else
        {
            var year = fileEditionDate.Year;
            var month = fileEditionDate.Month;
            var entryMonth = match.Groups["EffectiveDateMonth"].Value;
            var yearOffset = (entryMonth == "12" && month == 1) ? -1 : 0;
            effectiveDate = rowEffectiveDate + "/" + (year + yearOffset);
            valueDate = "";
        }

        {
            var accountEntry = new AccountEntry
            {
                AccountName = accountName,
                EntryType = entryType,
                EffectiveDate = DateTime.Parse(effectiveDate),
                ValueDate = (string.IsNullOrEmpty(valueDate)) ? null : DateTime.Parse(valueDate),
                Label = match.Groups["Label"].Value,
                Amount = GetAmount(match, isCardAccount, forSG)
            };

            System.Diagnostics.Debug.Print($"{accountEntry.AccountName}, {accountEntry.EntryType} : {accountEntry.EffectiveDate.ToShortDateString()}, {accountEntry.Amount}");

            accountEntries.Add(accountEntry);
        }
    }
}

decimal? GetAmount(Match match, bool isCardAccount, bool forSG)
{
    decimal? r = 0;
    var amountHasBeenFound = match.Groups["Amount"].Success;
    r = amountHasBeenFound ? (decimal.Parse(match.Groups["Amount"].Value.Replace(".", ""))) : (decimal?)null;
    if (r.HasValue && isCardAccount) r = -Math.Abs(r.Value);
    return r;
}

DateTime GetFileEditionDate(string fileName)
{
    var regex = new Regex(@"_(\d{8})\.");
    var match = regex.Match(fileName);

    if (match.Success && DateTime.TryParseExact(match.Groups[1].Value, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime date))
    {
        return date;
    }

    throw new ArgumentException();
}

// Export to CSV
var timeStamp = DateTime.Now.ToString("yyyyMMdd-hhmmss");
var exportFileName = $"Export des comptes - {timeStamp}.csv";
var exportPath = "\\\\SAPIENCE12\\Partage\\Pierre Wattenne\\Documents justificatifs\\Relevés de compte";
exportPath = "C:\\Users\\patrice.ongla.ext\\Downloads";
var exportFullFileName = Path.Combine(exportPath, exportFileName);
using (var writer = new StreamWriter(exportFullFileName))
using (var csv = new CsvWriter(writer, new CsvHelper.Configuration.CsvConfiguration(CultureInfo.CurrentCulture)
{
    Delimiter = ";"
}))
{
    csv.WriteRecords(accountEntries);
}
;

var msg = $"Fini. {accountEntries.Count} mouvements, Montant total : {accountEntries.Sum(d=> d.Amount)}. \nData exported in '{exportFullFileName}'";
System.Diagnostics.Debug.Print(msg);
Console.WriteLine(msg);

void SetSubPatterns(bool forSG, string entryType)
{
    if (forSG)
    {
        datePattern = (entryType == "CB") ? longDatePattern : veryLongDatePattern;
        amountPattern = amountPattern.Replace(@" ", @"\.");
    }
    else
    {
        datePattern = shortDatePattern;
        amountPattern = amountPattern.Replace(@"\.", @" ");
    }

}

string GetPattern(bool forSG, string entryType)
{
    var EffectiveDatePattern = $"(?'EffectiveDate'{datePattern.Replace("EntryMonth", "EffectiveDateMonth")})";
    var valueDatePattern = $"(?'ValueDate'{datePattern.Replace("EntryMonth", "ValueDateMonth")})";
    var labelPattern = "(?'Label'.+?)";


    string r;
    if (forSG)
    {
        var valueDatePart = (entryType == "CB") ? "" : valueDatePattern + @"\s";
        r = $@"{EffectiveDatePattern}\s{valueDatePart}{labelPattern}{spaceOrEndOfLinePattern}{amountPattern}";
    }
    else
    {
        datePattern = @"\d{2}/\d{2}";
        amountPattern = amountPattern.Replace(@"\.", @"\s");
        r = $@"{EffectiveDatePattern}\s{labelPattern}{spaceOrEndOfLinePattern}{amountPattern}";
    }

    return r;
}

string GetTextToParse(string allText, bool forSG, string entryType)
{
    int parsingStart, parsingEnd = 0;
    var startTextPattern = "";
    var endTextMark = "";
    if (forSG)
    {
        startTextPattern = (entryType == "CB") ? $@"(?:Opérations effectuées en France en euros :)" : $@"SOLDE PRÉCÉDENT AU {veryLongDatePattern}{spaceOrEndOfLinePattern}{amountPattern}";
        endTextMark = (entryType == "CB") ? "TOTAL NET DES OPÉRATIONS" : "TOTAUX DES MOUVEMENTS";
    }
    else
    {
        startTextPattern = $@"Ancien solde au {veryLongDatePattern}{spaceOrEndOfLinePattern}{amountPattern}";
        endTextMark = "Nouveau solde au";
    }

    Regex startRegex = new Regex(startTextPattern);
    var startMatch = startRegex.Match(allText);

    //var monText = "Ancien solde au 10/12/2019\r839,35";
    //startRegex = new Regex($@"Ancien solde au {longDatePattern}{spaceOrEndOfLinePattern}{amountPattern}");
    //startMatch = startRegex.Match(allText);

    parsingStart = startMatch.Index + startMatch.Length;
    parsingEnd = allText.IndexOf(endTextMark);

    var r = allText.Substring(parsingStart, parsingEnd - parsingStart);
    return r;
}

public class AccountEntry
{
    public String AccountName { get; set; }
    public String EntryType { get; set; }
    public DateTime EffectiveDate { get; set; }
    public DateTime? ValueDate { get; set; }
    public string Label { get; set; }
    public decimal? Amount { get; set; }
}
