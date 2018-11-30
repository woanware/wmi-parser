using CsvHelper;
using Fclp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace wmi
{
    /// <summary>
    /// 
    /// </summary>
    internal class ApplicationArguments
    {
        public string Output { get; set; }
        public string Input { get; set; }
    }

    class Program
    {
        #region Member Variables
        private static FluentCommandLineParser<ApplicationArguments> fclp;
        #endregion

        /// <summary>
        /// Application entry point
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (ProcessCommandLine(args) == false)
            {
                return;
            }

            if (CheckCommandLine() == false)
            {
                return;
            }

            string data = File.ReadAllText(fclp.Object.Input);

            Regex regexConsumer = new Regex(@"([\w_]*EventConsumer)\.Name=""([\w\s]*)""", RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var matchesConsumer = regexConsumer.Matches(data);

            Regex regexFilter = new Regex(@"_EventFilter\.Name=""([\w\s]*)""", RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var matchesFilter = regexFilter.Matches(data);

            List<Binding> bindings = new List<Binding>();
            for (int index = 0; index < matchesConsumer.Count; index++)
            {
                bindings.Add(new Binding(matchesConsumer[index].Groups[2].Value, matchesFilter[index].Groups[1].Value));
            }

            foreach (var b in bindings)
            {
                Regex regexEventConsumer = new Regex(@"\x00CommandLineEventConsumer\x00\x00(.*?)\x00.*?" + b.Name + "\x00\x00?([^\x00]*)?", RegexOptions.Multiline);

                var matches = regexEventConsumer.Matches(data);
                foreach (Match m in matches)
                {
                    b.Type = "CommandLineEventConsumer";
                    b.Arguments = m.Groups[1].Value;
                }

                regexEventConsumer = new Regex(@"(\w*EventConsumer)(.*?)(" + b.Name + @")(\x00\x00)([^\x00]*)(\x00\x00)([^\x00]*)", RegexOptions.Multiline);
                matches = regexEventConsumer.Matches(data);
                foreach (Match m in matches)
                {
                    b.Other = string.Format("{0} ~ {1} ~ {2} ~ {3}", m.Groups[1], m.Groups[3], m.Groups[5], m.Groups[7]);
                }

                regexEventConsumer = new Regex(@"(" + b.Filter + ")(\x00\x00)([^\x00]*)(\x00\x00)", RegexOptions.Multiline);
                matches = regexEventConsumer.Matches(data);
                foreach (Match m in matches)
                {
                    b.Query = m.Groups[3].Value;
                }
            }

            OutputToConsole(bindings);

            if (fclp.Object.Output != null)
            {
                OutputToFile(bindings);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static bool ProcessCommandLine(string[] args)
        {
            fclp = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            fclp.Setup(arg => arg.Input)
               .As('i')
               .Required()
               .WithDescription("Input file (OBJECTS.DATA)");

            fclp.Setup(arg => arg.Output)
                .As('o')
                .WithDescription("Output directory for analysis results");

            var header =
               $"\r\n{Assembly.GetExecutingAssembly().GetName().Name} v{Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
               "\r\n\r\nAuthor: Mark Woan / woanware (markwoan@gmail.com)" +
               "\r\nhttps://github.com/woanware/wmi-parser";

            // Sets up the parser to execute the callback when -? or --help is supplied
            fclp.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => Console.WriteLine(text));

            var result = fclp.Parse(args);

            if (result.HelpCalled)
            {
                return false;
            }

            if (result.HasErrors)
            {
                Console.WriteLine("");
                Console.WriteLine(result.ErrorText);
                fclp.HelpOption.ShowHelp(fclp.Options);
                return false;
            }

            Console.WriteLine(header);
            Console.WriteLine("");

            return true;
        }

        /// <summary>
        /// Performs some basic command line parameter checking
        /// </summary>
        /// <returns></returns>
        private static bool CheckCommandLine()
        {
            FileAttributes fa = File.GetAttributes(fclp.Object.Input);
            if ((fa & FileAttributes.Directory) == FileAttributes.Directory)
            {
                if (Directory.Exists(fclp.Object.Input) == false)
                {
                    Console.WriteLine("Input directory (-i) does not exist");
                    return false;
                }
            }
            else
            {
                if (File.Exists(fclp.Object.Input) == false)
                {
                    Console.WriteLine("Input file (-i) does not exist");
                    return false;
                }
            }

            if (fclp.Object.Output != null)
            {
                if (Directory.Exists(fclp.Object.Output) == false)
                {
                    Console.WriteLine("Output directory (-o) does not exist");
                    return false;
                }
            }            

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bindings"></param>
        private static void OutputToConsole(List<Binding> bindings)
        {
            // Output the data
            foreach (var b in bindings)
            {
                if ((b.Name.Contains("BVTConsumer") && b.Filter.Contains("BVTFilter")) || (b.Name.Contains("SCM Event Log Consumer") && b.Filter.Contains("SCM Event Log Filter")))
                {
                    Console.WriteLine("  {0}-{1} - (Common binding based on consumer and filter names,  possibly legitimate)", b.Name, b.Filter);
                }
                else
                {
                    Console.WriteLine("  {0}-{1}\n", b.Name, b.Filter);
                }

                if (b.Type == "CommandLineEventConsumer")
                {
                    Console.WriteLine("    Name: {0}", b.Name);
                    Console.WriteLine("    Type: {0}", "CommandLineEventConsumer");
                    Console.WriteLine("    Arguments: {0}", b.Arguments);
                }
                else
                {
                    Console.WriteLine("    Consumer: {0}", b.Other);
                }

                Console.WriteLine("\n    Filter:");
                Console.WriteLine("      Filter Name : {0}     ", b.Filter);
                Console.WriteLine("      Filter Query: {0}     ", b.Query);
                Console.WriteLine("");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bindings"></param>
        private static void OutputToFile(List<Binding> bindings)
        {
            using (FileStream fileStream = new FileStream(Path.Combine(fclp.Object.Output, "wmi-parser.tsv"), FileMode.Create, FileAccess.Write))
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            using (CsvWriter cw = new CsvHelper.CsvWriter(streamWriter))
            {
                cw.Configuration.Delimiter = "\t";
                // Write out the file headers
                cw.WriteField("Name");
                cw.WriteField("Type");
                cw.WriteField("Arguments");
                cw.WriteField("Filter Name");
                cw.WriteField("Filter Query");
                cw.NextRecord();

                foreach (var b in bindings)
                {
                    cw.WriteField(b.Name);
                    cw.WriteField(b.Type);
                    cw.WriteField(b.Arguments);
                    cw.WriteField(b.Filter);
                    cw.WriteField(b.Query);
                    cw.NextRecord(); ;
                }
            }
        }
    }
}
