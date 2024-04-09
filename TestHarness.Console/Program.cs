using ReadExcelFile.Lib;

namespace TestHarness.Console
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();

            program.Run();
        }

        private void Run()
        {
            ProcessExcelFile processExcelFile = new ProcessExcelFile();

            processExcelFile.ReadExcelFile();

            System.Console.ReadLine();
        }
    }
}
