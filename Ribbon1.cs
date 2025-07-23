using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

public partial class Ribbon1 : Office.IRibbonExtensibility
{
    private Office.IRibbonUI ribbon;

    public void OnCompareClick(Office.IRibbonControl control)
    {
        try
        {
            // Ask user to select two Excel files
            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls"
            };

            if (dialog.ShowDialog() != DialogResult.OK || dialog.FileNames.Length != 2)
            {
                MessageBox.Show("Please select exactly two Excel files.");
                return;
            }

            string file1 = dialog.FileNames[0];
            string file2 = dialog.FileNames[1];

            var stories1 = ReadStoriesFromFile(file1);
            var stories2 = ReadStoriesFromFile(file2);

            var results = CompareStories(stories1, stories2);

            OutputResultsToNewSheet(results);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error: " + ex.Message);
        }
    }

    private List<(string id, string text)> ReadStoriesFromFile(string path)
    {
        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        var workbook = excelApp.Workbooks.Open(path);
        var sheet = (Worksheet)workbook.Sheets[1];
        var range = sheet.UsedRange;

        var stories = new List<(string id, string text)>();

        for (int i = 2; i <= range.Rows.Count; i++) // assumes row 1 is headers
        {
            string id = Convert.ToString(((Range)range.Cells[i, 1]).Value2);
            string text = Convert.ToString(((Range)range.Cells[i, 2]).Value2);
            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(text))
                stories.Add((id, text));
        }

        workbook.Close(false);
        excelApp.Quit();
        return stories;
    }

    private List<(string id1, string text1, string id2, string text2, double similarity)> CompareStories(
        List<(string id, string text)> list1,
        List<(string id, string text)> list2)
    {
        var results = new List<(string, string, string, string, double)>();

        foreach (var (id1, text1) in list1)
        {
            foreach (var (id2, text2) in list2)
            {
                double similarity = ComputeCosineSimilarity(text1, text2);
                results.Add((id1, text1, id2, text2, similarity));
            }
        }

        return results.OrderByDescending(r => r.similarity).ToList();
    }

    private double ComputeCosineSimilarity(string text1, string text2)
    {
        var words1 = text1.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var words2 = text2.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries);

        var allWords = words1.Concat(words2).Distinct();

        var vec1 = allWords.Select(w => words1.Count(x => x == w)).ToArray();
        var vec2 = allWords.Select(w => words2.Count(x => x == w)).ToArray();

        double dot = 0, mag1 = 0, mag2 = 0;
        for (int i = 0; i < vec1.Length; i++)
        {
            dot += vec1[i] * vec2[i];
            mag1 += vec1[i] * vec1[i];
            mag2 += vec2[i] * vec2[i];
        }

        return (mag1 == 0 || mag2 == 0) ? 0 : dot / (Math.Sqrt(mag1) * Math.Sqrt(mag2));
    }

    private void OutputResultsToNewSheet(List<(string id1, string text1, string id2, string text2, double similarity)> results)
    {
        var excelApp = Globals.ThisAddIn.Application;
        var workbook = excelApp.ActiveWorkbook;
        var sheet = (Worksheet)workbook.Sheets.Add();
        sheet.Name = "Comparison Results";

        sheet.Cells[1, 1] = "ID 1";
        sheet.Cells[1, 2] = "Description 1";
        sheet.Cells[1, 3] = "ID 2";
        sheet.Cells[1, 4] = "Description 2";
        sheet.Cells[1, 5] = "Similarity (%)";

        int row = 2;
        foreach (var (id1, text1, id2, text2, sim) in results)
        {
            sheet.Cells[row, 1] = id1;
            sheet.Cells[row, 2] = text1;
            sheet.Cells[row, 3] = id2;
            sheet.Cells[row, 4] = text2;
            sheet.Cells[row, 5] = Math.Round(sim * 100, 2); // percent
            row++;
        }

        sheet.Columns.AutoFit();
        MessageBox.Show("Comparison complete. Results added to new sheet.");
    }

    public string GetCustomUI(string ribbonID) => null;
}
