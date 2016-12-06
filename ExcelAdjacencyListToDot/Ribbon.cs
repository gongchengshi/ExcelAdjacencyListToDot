using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAdjacencyListToDot
{
   public partial class Ribbon
   {
      private void RunBtn_Click(object sender, RibbonControlEventArgs e)
      {
         Run();
      }

      private void Run()
      {
         var dotString = ContructDot();
         var dotFile = WriteDotFile(dotString);

         //var image = CreateImageUsingApi(dotString);
         var image = CreateImageUsingCommand(dotFile);

         var imageSheet = FindImageSheet();

         Clipboard.SetImage(image);
         imageSheet.Paste(imageSheet.Cells[1, 1], "dotImage");

         image.Dispose();
      }

      private static string ContructDot()
      {
         Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet;

         int numBlankRows = 0;
         var dotFileContents = new StringBuilder();
         dotFileContents.Append("digraph g{\n");

         for (int r = 1; ; ++r)
         {
            var nodeLabel = ((thisSheet.Cells[r, 1] as Excel.Range).Value as string);

            if (string.IsNullOrWhiteSpace(nodeLabel))
            {
               ++numBlankRows;
               if (numBlankRows == 1)
               {
                  dotFileContents.Append("\n");
               }
               else if (numBlankRows == 2)
               {
                  break; // Stop after encountering 3 blank rows in a row
               }
               continue;
            }
            else
            {
               numBlankRows = 0;
            }

            var startNodeName = FormatNodeName(nodeLabel);
            dotFileContents.Append(string.Format("{0}[label=\"{1}\"];\n", startNodeName, nodeLabel));

            for (int c = 2; ; ++c)
            {
               var nodeName = ((thisSheet.Cells[r, c] as Excel.Range).Value as string);
               if (string.IsNullOrWhiteSpace(nodeName))
               {
                  break;
               }

               nodeName = FormatNodeName(nodeName);

               dotFileContents.Append(string.Format("{0}->{1}\n", startNodeName, nodeName));
            }
            dotFileContents.Append("\n");
         }

         dotFileContents.Append("}\n");

         var dotFileString = dotFileContents.ToString();

         return dotFileString;
      }

      private Excel.Worksheet FindImageSheet()
      {
         var app = Globals.ThisAddIn.Application;
         Excel.Worksheet thisSheet = app.ActiveSheet;

         Excel.Workbook wb = app.ActiveWorkbook;

         Excel.Worksheet imageSheet = null;

         for (int i = 1; i <= wb.Worksheets.Count; ++i)
         {
            if (thisSheet == (Excel.Worksheet)(wb.Worksheets[i]))
            {
               if (i < wb.Worksheets.Count)
               {
                  imageSheet = wb.Worksheets[i + 1];
               }
            }
         }

         if (imageSheet == null)
         {
            imageSheet = (Excel.Worksheet)(wb.Worksheets.Add());
         }

         return imageSheet;
      }

      private Image CreateImageUsingApi(string dotFile)
      {
         try
         {
            return GraphvizApi.Graphviz.RenderImage(dotFile, "dot", "bmp");
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message, "Error!");
            return null;
         }
      }

      private Image CreateImageUsingCommand(string dotFile)
      {
         var imageFile = Path.ChangeExtension(dotFile, ".bmp");

         File.Delete(imageFile);

         Process p = new Process();
         p.StartInfo.UseShellExecute = false;
         p.StartInfo.RedirectStandardOutput = true;

         const string GraphvizBin64 = @"C:\Program Files (x86)\Graphviz 2.28\bin\dot.exe";
         const string GraphvizBin32 = @"C:\Program Files\Graphviz 2.28\bin\dot.exe";

         var executablePath = File.Exists(GraphvizBin64) ? GraphvizBin64 : GraphvizBin32;

         p.StartInfo.FileName = executablePath;
         p.StartInfo.Arguments = string.Format("\"{0}\" -T bmp -o \"{1}\"", dotFile, imageFile);
         p.Start();
         p.WaitForExit();

         var image = Image.FromFile(imageFile);

         return image;
      }

      private string WriteDotFile(string dotFileString)
      {
         var app = Globals.ThisAddIn.Application;

         Excel.Workbook wb = app.ActiveWorkbook;

         var path = string.IsNullOrWhiteSpace(wb.Path) ? app.DefaultFilePath : wb.Path;

         var outFile = Path.Combine(path, Path.ChangeExtension(wb.Name, "dot"));

         using (var writer = new StreamWriter(outFile))
         {
            writer.Write(dotFileString);
         }

         return outFile;
      }

      private static string FormatNodeName(string nodeName)
      {
         var pattern = new Regex("[^a-zA-Z0-9_]+");
         return pattern.Replace(nodeName.ToLowerInvariant(), "_");
      }
   }
}
