using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private static string getBetween(string strSource, string strStart, string strEnd)
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = (Word.Document)Globals.ThisAddIn.Application.ActiveDocument;
            Word.Selection sel = (Word.Selection)Globals.ThisAddIn.Application.Selection;
            sel.Text.Trim();
            if (sel.Text.IndexOf("$") == 0)
            {
                System.Diagnostics.Debug.WriteLine("oi");
                File.WriteAllText("C:/oi.txt", sel.Text); //sempre tem erro aqui
                // Start the child process
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                // Redirect the output stream of the child process.
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.Arguments = "/C java -jar C:/mathtoweb.jar C:/oi.txt -rep -unicode -line -stdout";
                p.Start();
                // Do not wait for the child process to exit before
                // reading to the end of its redirected stream.
                // p.WaitForExit();
                // Read the output stream first and then wait.
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                string data = getBetween(output, "<math ", "</math>");
                if (data != "")
                {
                    data = "<math " + data + "</math>";
                    System.Windows.Forms.Clipboard.SetText(data);
                    sel.Paste();
                    sel.OMaths[1].Type = Word.WdOMathType.wdOMathInline;
                }
            }
            else
            {
                string texto = File.ReadAllText("C:/oi.txt");
                sel.Delete();
                System.Windows.Forms.Clipboard.SetText(texto);
                sel.Paste();
            }
        }
    }
}
