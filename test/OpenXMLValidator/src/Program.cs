namespace Abap2Xlsx.OpenXMLValidator
{
    using System;
    using System.Linq;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;

    /// <summary>
    /// Simple command line wrapper around OpenXMLValidator.
    /// See CommandLineArgs for possible options.
    /// </summary>
    class Program
    {
        /// <summary>
        /// Entry point
        /// </summary>
        /// <param name="args">The arguments</param>
        static void Main(string[] args)
        {
            CommandLineArgs arguments = new CommandLineArgs();
            try
            {
                arguments.Parse(args);
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(Resources.Usage);
                Environment.Exit(1);
            }

            if (arguments.ShowUsage)
            {
                Console.WriteLine(Resources.Usage);
                Environment.Exit(0);
            }

            string[] files = Directory.GetFiles(arguments.Directory, arguments.SearchPattern);
            if (!files.Any())
            {
                Console.WriteLine("No matching files found");
                Environment.Exit(1);
            }

            for (int i = 1; i <= files.Length; i++)
            {
                Console.WriteLine("Validating file {0} from {1}", i, files.Length);
                ValidateFile(files[i - 1]);
            }
        }

        private static void ValidateFile(string file)
        {
            ConsoleColor color = Console.ForegroundColor;

            try
            {
                Console.WriteLine("File name: {0}", file);

                OpenXmlValidator validator = new OpenXmlValidator();

                using (var doc = SpreadsheetDocument.Open(file, true))
                {
                    var errors = validator.Validate(doc);
                    if (errors.Any())
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Found {0} validation errors: ", errors.Count());

                        int count = 0;
                        foreach (ValidationErrorInfo error in errors)
                        {
                            count++;
                            Console.WriteLine("Error " + count);
                            Console.WriteLine("Part: " + error.Part.Uri);
                            Console.WriteLine("Description: " + error.Description);
                            Console.WriteLine("Path: " + error.Path.XPath);
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Document valid");
                    }

                    Console.WriteLine();
                }
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Exception occured while validating file: {0} - {1}",e.GetType().ToString() ,e.Message);
            }
            finally
            {
                Console.ForegroundColor = color;
            }
        }
    }
}
