using System;
using System.IO;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types.InputFiles;

namespace PDFile
{
    class Program
    {
        private static MemoryStream _downloadStream;
        private static PdfConventer _pdfConventer;

        private static readonly TelegramBotClient Bot
            = new TelegramBotClient("Your API KEY");


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
            var message = messageEventArgs.Message;
            var filePath = (await Bot.GetFileAsync(message.Document.FileId)).FilePath;
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(message.Document.FileName);
            var extension = Path.GetExtension(message.Document.FileName);
            _downloadStream = new MemoryStream();

            await Bot.DownloadFileAsync(filePath, _downloadStream);

            switch (extension)
            {
                case ".xls":
                case ".xlsx":
                    _pdfConventer = new XlsxToPdf();
                    break;
                case ".docx":
                    _pdfConventer = new DocxToPdf();
                    break;
            }

            SendMessage(message.Chat.Id, fileNameWithoutExtension);
        }

        static void SendMessage(long chatId, string fileNameWithoutExtension)
        {
            Bot.SendDocumentAsync(
                chatId,
                new InputOnlineFile(new MemoryStream(_pdfConventer.ToPdf(_downloadStream)))
                {
                    FileName = fileNameWithoutExtension + ".pdf"
                });
        }
    }
}
