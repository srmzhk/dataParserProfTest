using System.Text;
using System.Net.NetworkInformation;
using DataParserProfTest.Core;

class Program
{
    public static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.Clear();
        try
        {
            if (!IsInternetAvailable())
            {
                Console.WriteLine("Нет подключения к интернету.");
            }
            else
            {
                TestParser testParser = new TestParser();
                testParser.CheckTestsInterests();
                testParser.CheckTestsOrientation();
                testParser.CheckTestsInlination();
                testParser.CheckTestsBrigs();
                testParser.CheckTestsSocialType();
                testParser.CheckTestsThinkingType();
                testParser.CheckTestsProfType();
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }

    static bool IsInternetAvailable()
    {
        try
        {
            using (var ping = new Ping())
            {
                var reply = ping.Send("www.google.com");
                return reply.Status == IPStatus.Success;
            }
        }
        catch
        {
            return false;
        }
    }
}