using Aspose.Words;
using System;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.InlineQueryResults;
using Telegram.Bot.Types.ReplyMarkups;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot;
using Telegram.Bot.Types.InputFiles;
using System.IO;
using System.Net;
using System.Collections.Generic;
using Newtonsoft.Json;


namespace WordToPdf_Bot
{
    class Program
    {
        private static TelegramBotClient? Bot;

        public static async Task Main()
        {
            Bot = new TelegramBotClient("5098869164:AAG21ljCi79fYvDpj7e6HWEsi5dwx9RSqZw");

            User me = await Bot.GetMeAsync();
            Console.Title = me.Username ?? "WordToPdf Bot";
            using var cts = new CancellationTokenSource();

            ReceiverOptions receiverOptions = new() { AllowedUpdates = { } };
            Bot.StartReceiving(HandleUpdateAsync,
                               HandleErrorAsync,
                               receiverOptions,
                               cancellationToken: cts.Token);

            Console.WriteLine($"Start listening for @{me.Username}");
            Console.ReadLine();

            cts.Cancel();
        }

        public static async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
        {
            var handler = update.Type switch
            {
                UpdateType.Message => BotOnMessageReceived(botClient, update.Message!),
                UpdateType.EditedMessage => BotOnMessageReceived(botClient, update.EditedMessage!),
                _ => UnknownUpdateHandlerAsync(botClient, update)
            };

            try
            {
                await handler;
            }
            catch (Exception exception)
            {
                await HandleErrorAsync(botClient, exception, cancellationToken);
            }
        }

        private static async Task BotOnMessageReceived(ITelegramBotClient botClient, Message message)
        {
            Console.WriteLine($"Receive message type: {message.Type}");
            
            if(message.Text == "/help" || message.Text == "/start")
            {
                await botClient.SendTextMessageAsync(chatId: message.Chat.Id,
                                                                    text:
                                                                          "/help - Get help\n" +
                                                                          "Send the Word file(Doc/Docx) you want to convert to PDF.\n",
                                                                    replyMarkup: new ReplyKeyboardRemove());
                return;
            }

            else if ((message.Type != MessageType.Document))
                    {
                        return;
                    }


            var action = message.Text switch
            {
                //"/help" or "/start" => help(botClient, message),
                _ => wordToPdf(botClient, message)
            };
            Message sentMessage = await action;
            Console.WriteLine($"The message was sent with id: {sentMessage.MessageId}");


            static async Task<Message> help(ITelegramBotClient botClient, Message message)
            {


                return await botClient.SendTextMessageAsync(chatId: message.Chat.Id,
                                                            text:
                                                                  "/help - Get help\n" +
                                                                  "Send the Word file(Doc/Docx) you want to convert to PDF.\n",
                                                            replyMarkup: new ReplyKeyboardRemove());
            }



        }


        static async Task<Message> wordToPdf(ITelegramBotClient botClient, Message message)
        {


            //var rFile = message.Document;
            var fileId = message.Document.FileId;
            var fileName = message.Document.FileName;
            string[] fileNameArray = fileName.Split(".");

           

            if ((fileNameArray[1].ToUpper() == "DOC") || (fileNameArray[1].ToUpper() == "DOCX"))
            {

                await botClient.SendTextMessageAsync(chatId: message.Chat.Id,
                                                                text: "Please wait...");

                // Simulate longer running task
                await Task.Delay(5000);

                var webClient = new System.Net.WebClient();
                var json = webClient.DownloadString("https://api.telegram.org/bot5098869164:AAG21ljCi79fYvDpj7e6HWEsi5dwx9RSqZw/getFile?file_id=" + fileId);

                var jsonData = JsonConvert.DeserializeObject<Root>(json);

                string filePath = jsonData.result.file_path;
                string fullPath = "https://api.telegram.org/file/bot5098869164:AAG21ljCi79fYvDpj7e6HWEsi5dwx9RSqZw/" + filePath;



                //System.IO.File.WriteAllText("dsfsd",rFile);
                // Load the document from disk.
                Aspose.Words.Document doc = new Aspose.Words.Document(fullPath);
                // Save as PDF
                doc.Save("D:\\output.pdf");


                using (FileStream stream = System.IO.File.OpenRead("D:\\output.pdf"))
                {
                    InputOnlineFile inputOnlineFile = new InputOnlineFile(stream, "output.pdf");
                    return await botClient.SendDocumentAsync(
                        chatId: message.Chat.Id,
                        document: inputOnlineFile,
                        caption: "<b>Here is your PDF file</b>.",
                        parseMode: ParseMode.Html
                        );
                }
            }
            else
            {
                return await botClient.SendTextMessageAsync(chatId: message.Chat.Id,
                                                                text: "Make sure you have sent a word file and try again!!");
            
             }

            /*
            return await botClient.SendDocumentAsync(
            chatId: message.Chat.Id,
            document: "D:\\output.pdf",
            caption: "<b>Ara bird</b>. <i>Source</i>: <a href=\"https://pixabay.com\">Pixabay</a>",
            parseMode: ParseMode.Html);*/
        }


        private static Task UnknownUpdateHandlerAsync(ITelegramBotClient botClient, Update update)
        {
            Console.WriteLine($"Unknown update type: {update.Type}");
            return Task.CompletedTask;
        }



        public static Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
        {
            var ErrorMessage = exception switch
            {
                ApiRequestException apiRequestException => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
                _ => exception.ToString()
            };

            Console.WriteLine(ErrorMessage);
            return Task.CompletedTask;
        }



      

    }


    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
    public class Result
    {
        public string file_id { get; set; }
        public string file_unique_id { get; set; }
        public int file_size { get; set; }
        public string file_path { get; set; }
    }

    public class Root
    {
        public bool ok { get; set; }
        public Result result { get; set; }
    }



}
