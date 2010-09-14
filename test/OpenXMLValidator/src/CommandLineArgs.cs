namespace Abap2Xlsx.OpenXMLValidator
{
    using System;
    using System.IO;

    /// <summary>
    /// Command line arguments for validator
    /// </summary>
    public class CommandLineArgs
    {

        private const string DefaultDirectory = ".";
        private const string DefaultSearchPattern = "*.xlsx";

        public string Directory { get; private set; }
        public string SearchPattern { get; private set; }
        public bool ShowUsage { get; private set; }

        public CommandLineArgs()
        {
            this.Directory = DefaultDirectory;
            this.SearchPattern = DefaultSearchPattern;
        }

        public void Parse(string[] args)
        {
            for (int currentArg = 0; currentArg < args.Length; currentArg++)
            {
                string arg = args[currentArg];

                switch (arg)
                {
                    case "-d":
                    case "/d":
                        string dir = GetNextArg(args, currentArg);

                        DirectoryInfo di = new DirectoryInfo(dir);
                        if (!di.Exists)
                        {
                            throw new ArgumentException(string.Format("Unknown directory: {0}", dir));
                        }

                        this.Directory = dir;
                        currentArg++;

                        break;

                    case "-p":
                    case "/p":
                        this.SearchPattern = GetNextArg(args, currentArg);
                        currentArg++;

                        break;

                    case "-?":
                    case "/?":
                        this.ShowUsage = true;

                        break;

                    default:
                        throw new ArgumentException(string.Format("Unknown argument: {0}", arg));
                }
            }
        }

        private string GetNextArg(string[] args, int currentArg)
        {
            if (currentArg >= args.Length - 1)
            {
             throw new ArgumentException(string.Format("Missing value for argument: {0}", args[currentArg]));
            }

            return args[currentArg + 1];
        }
    }
}
