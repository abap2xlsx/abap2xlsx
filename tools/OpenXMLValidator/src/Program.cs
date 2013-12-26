namespace Abap2Xlsx.OpenXMLValidator
{
    using System;
    using System.Linq;
    using System.IO;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using System.Collections.Generic;

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

            string[] patterns = arguments.SearchPattern.Split('|');
            List<string> files = new List<string>();
            foreach (string pattern in patterns)
            {
                files.AddRange(Directory.GetFiles(arguments.Directory, pattern));
            }
            
            if (!files.Any())
            {
                Console.WriteLine("No matching files found");
                Environment.Exit(1);
            }

            int validFiles = 0;
            for (int i = 1; i <= files.Count; i++)
            {
                Console.WriteLine("Validating file {0} from {1}", i, files.Count);
                if (ValidateFile(files[i - 1]))
                {
                    validFiles++;
                } 
            }

            Console.WriteLine("Files checked  - {0}", files.Count);
            Console.WriteLine("Valid  files   - {0}", validFiles);
            Console.WriteLine("Invalid  files - {0}", files.Count - validFiles);
        }

        /// <summary>
        /// Validates the file and prints result to console
        /// </summary>
        /// <param name="file">Path to the file</param>
        private static bool ValidateFile(string file)
        {
            ConsoleColor color = Console.ForegroundColor;
            bool isValid = false;

            try
            {
                Console.WriteLine("File name: {0}", file);

                OpenXmlValidator validator = new OpenXmlValidator();

                using (var doc = GetOpenXmlPackage(file))
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
                        isValid = true;
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

            return isValid;
        }

        /// <summary>
        /// Returns OpenXmlPackage instance for a file.
        /// .docx, .xlsx and .pptx files are supported
        /// </summary>
        /// <param name="file">Path to file</param>
        /// <returns>OpenXmlPackage instance</returns>
        private static OpenXmlPackage GetOpenXmlPackage(string file)
        {
            FileInfo fi = new FileInfo(file); 
            
            switch (fi.Extension.ToLowerInvariant())
            {
                case ".xlsx":
                    return SpreadsheetDocument.Open(file, true);
                case ".docx": 
                    return WordprocessingDocument.Open(file, true);
                case ".pptx": 
                    return PresentationDocument.Open(file, true);
                default:
                    throw new ArgumentException(string.Format("Unknown file extension {0}", fi.Extension), "file");
            }
            
        }
    }
}
