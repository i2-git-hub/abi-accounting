using System;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using i2.accounting.sdk;

namespace abi_automation
{
    class Program 
    {
        static string pdf_file_location = @"";
        static string[] file_entries;
        static string final_file_name;
        static file_paths _file_paths;
        static Vendor_Functions _vendor_function;
        static void Main(string[] args)
        {
            Console.WriteLine("Reading File Directory....");
            _file_paths = new file_paths();
            process_directory(_file_paths._directiory_path);
            progress_bar();
            Console.WriteLine("Click Enter to Close");
            Console.ReadKey();
        }

        static void process_directory(string _path)
        {
            file_entries = Directory.GetFiles(_path);
            foreach(string fileName in file_entries)
            {
                try
                {
                    processFile(fileName);
                }
                catch(Exception er)
                {
                    Console.WriteLine(er);
                }
            }
        }

        static void processFile(string file)
        {
            pdf_file_location = file;
            on_start(pdf_file_location, _file_paths._raw_pdf_data);
            identify_vendor(_file_paths._raw_pdf_data);
        }

        static void on_start(string pdfFile, string textFile)
        {
            pdfFile = pdfFile.Replace("\r\n", string.Empty);
            using (BitMiracle.Docotic.Pdf.PdfDocument doc = new BitMiracle.Docotic.Pdf.PdfDocument(pdfFile))
            {
                string text = doc.GetTextWithFormatting();
                File.WriteAllText(textFile, text);
            }
        }

        static void identify_vendor(string textFile)
        {
            _vendor_function = new Vendor_Functions();
            bool heilind_bool_value = _vendor_function.vendor_identifying_function(textFile, "Heilind Electronics, Inc.");
            if(heilind_bool_value == true)
            {
                final_file_name = _vendor_function.find_invoice_using_skip_method_with_substring_values(textFile, 19, 1, 0, 6, "_heilind_");
                Double invoice_total_with_tariff = Convert.ToDouble(_vendor_function.find_invoice_total_using_keyword_split_array_indexing(textFile, "NET SALES", 'D', 1)) + Convert.ToDouble(_vendor_function.find_tariff_value(textFile));
                final_file_name = final_file_name + Convert.ToString(invoice_total_with_tariff);
                rename_file(final_file_name);
            }
        }

        static void rename_file(string final_file_path)
        {
            final_file_name = _file_paths._directiory_path + final_file_path + ".pdf";
            File.Move(pdf_file_location, final_file_name);
            final_file_name = "";
        }


        static void progress_bar()
        {
            using (var progress = new ProgressBar())
            {
                for (int i = 0; i <= 100; i++)
                {
                    progress.Report((double)i / 100);
                    Thread.Sleep(20);
                }
            }
            Console.WriteLine("Done");
        }

        public class ProgressBar : IDisposable, IProgress<double>
        {
            private const int blockCount = 10;
            private readonly TimeSpan animationInterval = TimeSpan.FromSeconds(1.0 / 8);
            private const string animation = @"|/-\";

            private readonly Timer timer;

            private double currentProgress = 0;
            private string currentText = string.Empty;
            private bool disposed = false;
            private int animationIndex = 0;

            public ProgressBar()
            {
                timer = new Timer(TimerHandler);

                // A progress bar is only for temporary display in a console window.
                // If the console output is redirected to a file, draw nothing.
                // Otherwise, we'll end up with a lot of garbage in the target file.
                if (!Console.IsOutputRedirected)
                {
                    ResetTimer();
                }
            }

            public void Report(double value)
            {
                // Make sure value is in [0..1] range
                value = Math.Max(0, Math.Min(1, value));
                Interlocked.Exchange(ref currentProgress, value);
            }

            private void TimerHandler(object state)
            {
                lock (timer)
                {
                    if (disposed) return;

                    int progressBlockCount = (int)(currentProgress * blockCount);
                    int percent = (int)(currentProgress * 100);
                    string text = string.Format("[{0}{1}] {2,3}% {3}",
                        new string('#', progressBlockCount), new string('-', blockCount - progressBlockCount),
                        percent,
                        animation[animationIndex++ % animation.Length]);
                    UpdateText(text);

                    ResetTimer();
                }
            }

            private void UpdateText(string text)
            {
                // Get length of common portion
                int commonPrefixLength = 0;
                int commonLength = Math.Min(currentText.Length, text.Length);
                while (commonPrefixLength < commonLength && text[commonPrefixLength] == currentText[commonPrefixLength])
                {
                    commonPrefixLength++;
                }

                // Backtrack to the first differing character
                StringBuilder outputBuilder = new StringBuilder();
                outputBuilder.Append('\b', currentText.Length - commonPrefixLength);

                // Output new suffix
                outputBuilder.Append(text.Substring(commonPrefixLength));

                // If the new text is shorter than the old one: delete overlapping characters
                int overlapCount = currentText.Length - text.Length;
                if (overlapCount > 0)
                {
                    outputBuilder.Append(' ', overlapCount);
                    outputBuilder.Append('\b', overlapCount);
                }

                Console.Write(outputBuilder);
                currentText = text;
            }

            private void ResetTimer()
            {
                timer.Change(animationInterval, TimeSpan.FromMilliseconds(-1));
            }

            public void Dispose()
            {
                lock (timer)
                {
                    disposed = true;
                    UpdateText(string.Empty);
                }
            }
        }
    }
}
