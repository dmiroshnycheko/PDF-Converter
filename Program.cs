using System;
using System.IO;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.InputFiles;

namespace PDFile
{
    class Program
    {
        private static readonly TelegramBotClient Bot
            = new TelegramBotClient("API KEY");


        static void Main(string[] args)
        {
            Bot.OnMessage += BotOnMessageReceived;
            var me = Bot.GetMeAsync().Result;
            Console.Title = me.Username;
            Bot.StartReceiving();
            Console.WriteLine($"Start listening for @{me.Username}");
            Console.ReadLine();
            Bot.StopReceiving();
        }


        private static async void BotOnMessageReceived(object sender, MessageEventArgs messageEventArgs)
        {
            try
            {
                await ProcessMessage(messageEventArgs);
            }
            catch (Exception e)
            {
                await SendMessage(messageEventArgs.Message.Chat.Id, e.Message);
            }
        }

        private static async Task ProcessMessage(MessageEventArgs messageEventArgs)
        {
            PdfConventer _pdfConventer;
            MemoryStream _downloadStream = new MemoryStream();
            var message = messageEventArgs.Message;
            if (message.Document == null)
                throw new Exception("Send the document with the extension .docx or .xls or .xlsx ");
            var filePath = (await Bot.GetFileAsync(message.Document.FileId)).FilePath;
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(message.Document.FileName);
            var extension = Path.GetExtension(message.Document.FileName);
            await Bot.DownloadFileAsync(filePath, _downloadStream);
            switch (extension)
            {
                case ".xls":
                case ".xlsx":

                    _pdfConventer = new XlsxToPdf();
                    Console.WriteLine(_pdfConventer.GetType().ToString());
                    break;
                case ".docx":
                    _pdfConventer = new DocxToPdf();
                    Console.WriteLine(_pdfConventer.GetType().ToString());
                    break;
                default:
                    throw new Exception("There is no such format. Supported Formats:docx,xls,xlsx ");
            }

            await SendFile(message.Chat.Id, fileNameWithoutExtension, _pdfConventer, _downloadStream);
        }

        static async Task SendMessage(long chatId, string text)
        {
            await Bot.SendTextMessageAsync(chatId, text);
        }

        static async Task SendFile(long chatId, string fileNameWithoutExtension, PdfConventer pdf, MemoryStream ds)
        {
            Console.WriteLine("...");
            Console.WriteLine(fileNameWithoutExtension);
            var a = Bot.SendDocumentAsync(
                chatId,
                new InputOnlineFile(new MemoryStream(pdf.ToPdf(ds)))
                {
                    FileName = fileNameWithoutExtension + ".pdf"
                }).Result;

            Console.WriteLine("!!!");
        }
    }
}
