using System;
using System.IO;
using PdfiumViewer;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using CommandLine;
using CommandLine.Text;
using System.Drawing.Printing;
using System.Linq;
using System.Printing;
using System.Management;

namespace PDFPrinter
{
    class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(
                    (options) =>
                    {
                        var print_settings = new PrinterSettings();
                        var printers = GetAllPrinters();

                        bool show_printer_list = false;
                        if (options.PrinterName == null)
                        {
                            show_printer_list = true;
                        } else {
                            if (!printers.Contains(options.PrinterName))
                            {
                                Console.WriteLine($"Printer name is invalid ({options.PrinterName})");
                                show_printer_list = true;
                            } else
                            {
                                print_settings.PrinterName = options.PrinterName;
                            }
                        }

                        if (show_printer_list) {
                            Console.WriteLine($"Available printers are:");
                            foreach (var printer in printers)
                            {
                                Console.WriteLine(printer);
                            }
                        }


                        // Create the printer settings, optionally specifying the printer name
                        if (print_settings.IsValid)
                        {
                            try {

                                // Load the PDF document into memory
                                PdfDocument document = null;
                                try {
                                    document = PdfDocument.Load(options.FileName);

                                }
                                catch (FileNotFoundException)
                                {
                                    Console.WriteLine("The file could not be found.");
                                    throw;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("An error occurred while loading the document: " + ex.Message);
                                    throw;
                                }

                                // Print the document
                                var print_doc = document.CreatePrintDocument();
                                print_doc.PrinterSettings.PrinterName = options.PrinterName;

                                // Get the page size
                                if (options.PageSize != null)
                                {
                                    Console.WriteLine("Page size provided");

                                    var numbers = options.PageSize.Split('x');
                                    decimal width = decimal.Parse(numbers[0]);
                                    decimal height = decimal.Parse(numbers[1]);

                                    // Convert to hundredths of an inch
                                    width *= 100 / 25.4m;
                                    height *= 100 / 25.4m;

                                    // Set the page size
                                    /*Console.WriteLine("Paper sizes available:");
                                    foreach (var size in print_doc.PrinterSettings.PaperSizes)
                                    {
                                        Console.WriteLine(size.ToString());
                                    }*/
                                    print_doc.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                                    print_doc.DefaultPageSettings.PaperSize = new PaperSize("Custom size", (int)width, (int)height);

                                    Console.WriteLine($"Printing at {(int)width} x {(int)height}");
                                }


                                // Number of copies
                                if (options.NumCopies < 1)
                                {
                                    throw new ArgumentException($"Cannot print less than one copy. Number provided was {options.NumCopies}");
                                } else if (options.NumCopies > 1)
                                {
                                    Console.WriteLine($"Printing {options.NumCopies} copies");
                                    print_doc.PrinterSettings.Copies = options.NumCopies;

                                    // If printing more than one copy, check for collation setting provided
                                    if (options.Collate != null)
                                    {
                                        print_doc.PrinterSettings.Collate = options.Collate.Value;
                                    }
                                }

                                // Colour printing
                                if (options.IsColour != null)
                                {
                                    if (options.IsColour.Value)
                                    {
                                        if (print_doc.PrinterSettings.SupportsColor)
                                        {
                                            print_doc.DefaultPageSettings.Color = true;
                                        } else
                                        {
                                            Console.WriteLine("Warning: Colour printing requested but printer does not support colour printing.");
                                        }
                                    } else
                                    {
                                        print_doc.DefaultPageSettings.Color = false;
                                    }
                                }


                                // Duplex settings
                                if (options.Duplex != null)
                                {
                                    if (options.Duplex == Options.DUPLEX_SINGLE)
                                    {
                                        if (print_doc.PrinterSettings.CanDuplex)
                                        {
                                            print_doc.PrinterSettings.Duplex = Duplex.Simplex;
                                        }

                                    } else {
                                        if (print_doc.PrinterSettings.CanDuplex)
                                        {
                                            switch (options.Duplex)
                                            {
                                                case Options.DUPLEX_PORTRAIT:
                                                    print_doc.PrinterSettings.Duplex = Duplex.Vertical;
                                                    break;

                                                case Options.DUPLEX_LANDSCAPE:
                                                    print_doc.PrinterSettings.Duplex = Duplex.Horizontal;
                                                    break;

                                                default:
                                                    throw new ArgumentException($"Invalid duplex setting. Valid values are {Options.DUPLEX_SINGLE}, {Options.DUPLEX_PORTRAIT} or {Options.DUPLEX_LANDSCAPE}");
                                            }
                                        } else
                                        {
                                            Console.WriteLine($"Warning: Duplex setting provided ({options.Duplex}) but printer does not support duplex printing.");
                                        }
                                    }
                                }


                                // Set the tray to print to, if specified
                                if (options.Tray != null)
                                {

                                    bool tray_set = false;
                                    foreach (PaperSource source in print_doc.PrinterSettings.PaperSources)
                                    {
                                        if (source.SourceName.Trim().Equals(options.Tray))
                                        {
                                            print_doc.DefaultPageSettings.PaperSource = source;
                                            tray_set = true;
                                            break;
                                        }
                                    }

                                    if (!tray_set)
                                    {
                                        Console.WriteLine($"Paper source \"{options.Tray}\" is invalid.");
                                        Console.WriteLine($"Available paper sources:");
                                        foreach (PaperSource source in print_doc.PrinterSettings.PaperSources)
                                        {
                                            Console.WriteLine(source.SourceName.Trim());
                                        }

                                        return;
                                    }

                                }


                                Console.Write($"Printing on {print_settings.PrinterName}...");
                                try {
                                    print_doc.BeginPrint += PrintStarted;
                                    print_doc.EndPrint += PrintEnded;
                                    print_doc.Print();

                                } catch (Exception ex)
                                {
                                    Console.WriteLine($"Failed");
                                    Console.WriteLine($"An error occurred during printing: {ex.Message}");
                                    return;
                                }
                                Console.WriteLine($"Done");

                            } catch (Exception ex)
                            {
                                Console.WriteLine($"An error occurred during printing: {ex.Message}/n{ex.StackTrace}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Print settings are invalid: {print_settings.PrinterName}.");

                            if (options.IsInteractive)
                            {
                                Console.WriteLine("Press any key to quit");
                                Console.Read();
                            }

                            return;
                        }
                    })
                .WithNotParsed((errors) =>
            {
                Console.WriteLine($"Error: Script arguments are invalid.");
                Console.WriteLine($"Errors: " + errors);
            });
        }

        private static void PrintStarted(object sender, PrintEventArgs e)
        {
            Console.WriteLine("Starting printing...");
        }

        private static void PrintEnded(object sender, PrintEventArgs e)
        {
            Console.WriteLine("Finished printing");
        }

        private static List<string> GetAllPrinters()
        {
            var printer_names = new List<string>();

            ManagementScope objScope = new ManagementScope(ManagementPath.DefaultPath); //For the local Access
            objScope.Connect();

            SelectQuery selectQuery = new SelectQuery()
            {
                QueryString = "Select * from win32_Printer"
            };
            ManagementObjectSearcher MOS = new ManagementObjectSearcher(objScope, selectQuery);
            ManagementObjectCollection MOC = MOS.Get();
            foreach (ManagementObject mo in MOC)
            {
                printer_names.Add(mo["Name"].ToString());
            }

            return printer_names;
        }
    }



    // Define a class to receive parsed values
    class Options
    {
        public const string DUPLEX_SINGLE = "single";
        public const string DUPLEX_PORTRAIT = "portrait";
        public const string DUPLEX_LANDSCAPE = "landscape";

        [Option('p', "printerName", 
          HelpText = "Name of the printer to print the document to.",
            Required =true)]
        public string PrinterName { get; set; }

        /*[Option('o', "orientation",
          HelpText = "Orientation (Landscape/Portrait/Auto)")]
        public string Orientation { get; set; }*/

        [Option('s', "pageSize",
          HelpText = "Paper size to print to.")]
        public string PageSize { get; set; }

        [Option('i', "interactive",
          HelpText = "Run the program in interactive mode.")]
        public bool IsInteractive { get; set; }

        [Option('t', "tray",
          HelpText = "Paper source to print from.")]
        public string Tray { get; set; }


        [Option('f', "fileName",
          HelpText = "File path of the document to print.",
            Required =true)]
        public string FileName { get; set; }

        [Option('c', "colour",
          HelpText = "Whether or not to print in colour.")]
        public bool? IsColour { get; set; } = null;

        [Option('n', "copies",
          HelpText = "Number of copies to print.")]
        public short NumCopies { get; set; } = 1;

        [Option('x', "collate",
          HelpText = "Whether or not to collate copies.")]
        public bool? Collate { get; set; } = null;

        [Option('d', "duplex",
          HelpText = "Whether or not to print double sided.")]
        public string Duplex { get; set; } = DUPLEX_SINGLE;
    }
}
