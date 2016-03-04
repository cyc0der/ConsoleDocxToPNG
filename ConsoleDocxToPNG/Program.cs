using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Task2 = System.Threading.Tasks.Task;
using System.Drawing;

namespace ConsoleDocxToPNG
{
    class Program
    {
        const decimal PNGHeight = 5551;
        const decimal PNGWidth = 2550;

        static void Main(string[] args)
        {

            if (File.Exists("DocxToPNG_error.log"))
            {
                File.Delete("DocxToPNG_error.log");
            }

            if (File.Exists("DocxToPNG_error_fileList.log"))
            {
                File.Delete("DocxToPNG_error_fileList.log");
            }

            string[] list = Directory.GetFiles(Environment.CurrentDirectory + "/docx", "*.docx", System.IO.SearchOption.AllDirectories);



            foreach (string docx in list)
            {

                Application app = new Application();

                app.Visible = false;

                convert(docx, app,1);

            }

        }

        public static void convert(string docxStr, Application app, int pageNumStart)
        {

            FileInfo docx = new FileInfo(docxStr);

            string error;
            string destFolder = docx.FullName.Replace(".docx", "_png");

            if (!Directory.Exists(destFolder))
            {
                try
                {
                    Directory.CreateDirectory(destFolder);

                    Console.WriteLine(" - *_png directory Created.");
                }
                catch (Exception e)
                {
                    Program.moveToAnotherList(docxStr);
                    error = " *** can't create *_png directory\r\nPath:" + destFolder + "\r\nException:" + e.Message;
                    File.AppendAllText("DocxToPNG_error.log", error);
                    Console.WriteLine(error);
                    return;
                }
            }
            else
            {
                Console.WriteLine(" - *_png directory already exists.\r\n");
            }

            Document doc = new Document();

            try
            {
                doc = app.Documents.Open(docx.FullName);

                Console.WriteLine(" - docx file " + docx.Name + " opened.\r\n");
            }
            catch (Exception er)
            {
                Program.moveToAnotherList(docxStr);
                error = " *** can't open docx file\r\nFile:" + docx.FullName + "\r\nException:" + er.Message;
                File.AppendAllText("DocxToPNG_error.log", error);
                File.AppendAllText("DocxToPNG_error_fileList.log", docx.FullName + "\r\n");
                Console.WriteLine(error);
            }


            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;

            Console.WriteLine(" - start looping pages...\r\n");
            //Opens the word document and fetch each page and converts to image
            foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = pageNumStart; i <= pane.Pages.Count; i++)
                    {
                        Console.WriteLine(" - fetching page(" + i.ToString() + ") data\r\n");
                        Microsoft.Office.Interop.Word.Page page = null;
                        bool populated = false;
                        while (!populated)
                        {
                            try
                            {
                                // This !@#$ variable won't always be ready to spill its pages. If you step through
                                // the code, it will always work.  If you just execute it, it will crash.  So what
                                // I am doing is letting the code catch up a little by letting the thread sleep
                                // for a microsecond.  The second time around, this variable should populate ok.
                                page = pane.Pages[i];
                                populated = true;
                            }
                            catch (COMException ex)
                            {
                                Thread.Sleep(1);
                            }
                        }
                        var bits = page.EnhMetaFileBits;
                        var pngTarget = destFolder + "\\" + i.ToString() + ".png";

                        try
                        {
                            Console.WriteLine(" - getting from MemmoryStream page(" + i.ToString() + ")\r\n");
                            using (var ms = new MemoryStream((byte[])(bits)))
                            {
                                Image image = Image.FromStream(ms);
                                Bitmap myBitmap = new Bitmap(image, new Size(Convert.ToInt32(PNGWidth), Convert.ToInt32(PNGHeight)));
                                myBitmap.Save(pngTarget, System.Drawing.Imaging.ImageFormat.Png);
                                Console.WriteLine(" - PNG saved.\r\n");
                            }
                        }
                        catch (System.Exception ex)
                        {

                            if (app != null)
                            {
                                app.Quit(false, Type.Missing, Type.Missing);
                                Marshal.ReleaseComObject(app);
                                app = null;
                            }

                            app = new Application();

                            app.Visible = false;

                            Console.WriteLine(docxStr);

                            convert(docxStr, app, i);


                            return;

                               /* Program.moveToAnotherList(docxStr);
                                error = " *** can't save PNG page(" + i.ToString() + ")\r\nFile:" + docx.FullName + "\r\nException:" + ex.Message + ex.Source;
                                File.AppendAllText("DocxToPNG_error.log", error);
                                File.AppendAllText("DocxToPNG_error_fileList.log", docx.FullName + "\r\n");
                                Console.WriteLine(error);*/
                         }


                        
                    }
                }
            }

            doc.Close(false, Type.Missing, Type.Missing);


            Marshal.ReleaseComObject(doc);

            Console.WriteLine(" - docx closed.\r\n\r\n");



            if (app != null)
            {
                app.Quit(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(app);
                app = null;
            }

        }


        public static void moveToAnotherList(string path)
        {
            string destPath = Path.GetDirectoryName(path).Replace("\\docx\\", "\\docx-ConsoleDocxToPNG-failed\\");

            try
            {
                Directory.CreateDirectory(destPath);

                try
                {
                    File.Copy(path, destPath + "\\" + Path.GetFileName(path));

                    try
                    {
                        File.Delete(path);
						File.Delete(path);
                    }
                    catch (Exception e)
                    {
                        string error = "###  Can't delete moved damaged file \r\nPath:" + path + "\r\nException:" + e.Message;
                        File.AppendAllText("DocxToPNG_error.log", error);
                        Console.WriteLine(error);
                        return;
                    }
                }
                catch (Exception er)
                {
                    string error = "###  Can't  copy damaged file \r\nPath:" + path + "\r\nException:" + er.Message;
                    File.AppendAllText("DocxToPNG_error.log", error);
                    Console.WriteLine(error);
                    return;
                }
            }
            catch (Exception q)
            {
                string error = "###  Can't create directory for copy damaged file \r\nPath:" + destPath + "\r\nException:" + q.Message;
                File.AppendAllText("DocxToPNG_error.log", error);
                Console.WriteLine(error);
                return;
            }
        }
    }
}
