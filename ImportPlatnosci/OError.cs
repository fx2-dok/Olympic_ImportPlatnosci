
namespace ImportPlatnosci
{
    public class OError
    {
        public string Line { get; set; }
        public string ErrorMessage { get; set; }
        public string FileName { get; set; }
        public Payment Payment { get; set; }

        public OError(string line, string message, string fileName)
        {
            Line = line;
            ErrorMessage = message;
            FileName = fileName;
        }

        public OError(string line, string message, string fileName, Payment payment)
        {
            Line = line;
            ErrorMessage = message;
            FileName = fileName;
            Payment = payment;
        }
    }
}
