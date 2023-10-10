using System;
using VkNet;
using VkNet.Enums.SafetyEnums;
using VkNet.Model;
using VkNet.Model.RequestParams;
using System.Net;
using System.Threading;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
//Install-Package VkNet -Version 1.42.0
namespace ConsoleApp1
{
    class Program
    {
        public static VkApi api = new VkApi();
        static public void fuuu()
        {
            try
            {
                var s = api.Groups.GetLongPollServer(207499199);
                var poll = api.Groups.GetBotsLongPollHistory(
                   new BotsLongPollHistoryParams()
                   { Server = s.Server, Ts = s.Ts, Key = s.Key, Wait = 25 });
                if (poll?.Updates != null)
                {

                    foreach (var a in poll.Updates)
                    {
                        if (a.Type == GroupUpdateType.MessageNew)
                        {
                            string userMessage = a.Message.Text.ToLower();
                            long? userID = a.Message.PeerId;
                            e(userMessage, userID);
                        }
                    }
                }
            }
            catch { }

        }
        static void Main(string[] args)
        {

            try { 




            DirectoryInfo dirInfo = new DirectoryInfo(@"C://test//");
            if (!dirInfo.Exists)
            {
                Directory.CreateDirectory("C://test//");
            }
            DirectoryInfo dirInfo1 = new DirectoryInfo(@"C://test1//");
            if (!dirInfo1.Exists)
            {
                Directory.CreateDirectory("C://test1//");
            }
            FileInfo fi1 = new FileInfo(@"C://test1//Name.txt");
            if (!fi1.Exists)
            {
                fi1.CreateText().Close();
            }
            Class1 class1 = new Class1();
            class1.asincVoid();
                api.Authorize(new ApiAuthParams() { AccessToken = "31429b2b19826235bc53757b1d9d9b42736aba751e51646897833788d265c865c5825b813f974ac502d55" });
                while (true) // Бесконечный цикл, получение обновлений
                {
                    //
                    //
                    //
                    //
                    //Сделай по дням недели
                    //
                    //
                    //
                    //
                    //
                    fuuu();

                }
            }
            catch
            {
                Console.WriteLine("???");
                Thread.Sleep(300000);
                Main(args);
            }
        }
        public static void e(string message, long? userID)
        {
            try
            {
                if (!message.StartsWith("?"))
                {
                    if (Save.Safe1 == false) SendMessage("Привет.\n Наконец то добавлена функция рассылки. \n Чтоб получать рассылку при обновлении, добавьте ? преред любой из команд \n Пример: ?1292 или ?!Иванов \n Комманду нужно ввести 1 раз и когда обновится расписание, оно сразу придет вам сообщением \n Чтоб изменить данные для расылки повторите команду с ? \n Чтоб отключить рассылку напишите ?отменить", userID);
                }
                Class1 class1 = new Class1();
                //if (message == "привет") SendMessage("Приветик", userID);
                if (message == "гений евгеньевич") SendMessage("Леша, ты заебал моего бота засерать", userID);
                if (message == "начать") SendMessage(" Этот бот предназначен для выдачи расписания для Серпуховского Колледжа." +
                    " \n\n Чтоб получить расписани для 1 корпуса напишите номер группы. " +
                    " \n\n Чтоб получить расписани для преподавателей 1 корпуса напишите !фамилия_преподавателя. " +
                    " \n\n Чтоб получать рассылку при обновлении добавьте ? преред любой из команд. " +
                    "\n\n Для связи с разработчиком есть беседа в закрепе группы.", userID);
                if (message.StartsWith("?"))
                {
                    message = message.Replace("?", "");
                    Save.Safe = true;
                }
                if (Save.Safe)
                {
                    Class1 oo = new Class1();
                    oo.Sqve(message, userID);
                }
                int num;
                bool isNum = int.TryParse(message, out num);
                if (isNum)
                {
                    Class2.Gug = message;
                    class1.file();
                    geg(message, userID);
                }
                if (message == "сайт") SendMessage("http://serp-koll.ru/index.php/pages/raspisanie-urokov", userID);
                if (message.StartsWith("!"))
                {
                    message = message.Replace("!", "");
                    Classr.fam = message;
                    class1.filel();

                    gegr(message, userID);
                }
            }
            catch { }

            //if (message == "Фамилия")
            //{
            //    await Task.Run(() => rur(message, userID));
            //}
        }
        static public void gegr(string message, long? userID)
        {
            string[] zvon = new string[] { "08:45 - 10:15", "10:25 - 11:55", "12:20 - 13:50", "14:00 - 15:30", "15:40 - 17:10" };
            try
            {
                string[] mas = Class2.Mas;
                for (int i = 6; i > 1; i--)
                {
                    if (mas[i] == null || mas[i] == "")
                    {
                        mas[i] = "r";
                    }
                    else
                    {
                        break;
                    }
                }
                bool k = false;
                string j = "";
                for (int i = 0; i < 6; i++)
                {
                    if (mas[i] == "r")
                    {
                    }
                    else
                    {
                        if (i == 0)
                        {
                            //SendMessage($"{mas[i]}", userID);
                            j += $"{mas[i]}\n\n";
                        }
                        else
                        {
                            if (mas[i] != "" && mas[i] != null)
                            {
                                k = mas[i].ToLower().Contains("практика");
                            }
                                if (k)
                                {
                                    j += $"{mas[i]}\n\n";
                                    break;

                                }
                                if (mas[i] == "" || mas[i] == null)
                                    j += $"{i} пара - нет пары \n[{ zvon[i-1]}]\n\n";
                                else
                                    j += $"{i} пара - {mas[i]} \n[{ zvon[i-1]}]\n\n";

                           

                        }
                    }
                }

                SendMessage(j, userID);
            }
            catch { }
        }
        public static void rur(string message, long? userID)
        {
        }
        static public void geg(string message, long? userID)
        {
            string[] zvon = new string[] { "08:45 - 10:15", "10:25 - 11:55", "12:20 - 13:50", "14:00 - 15:30", "15:40 - 17:10"};





            try
            {
                string[] mas = Class2.Mas;
                for (int i = 13; i > 1; i--)
                {
                    if (mas[i] == null || mas[i] == "")
                    {
                        mas[i] = "r";
                    }
                    else
                    {
                        break;
                    }
                }
                bool k = false;
                string j = "";
                for (int i = 0; i < 14; i++)
                {

                    if (i == 1 || i == 0)
                    {
                        //SendMessage($"{mas[i]}", userID);
                        j += $"{mas[i]}\n\n";
                    }
                    else
                    {
                        if ((i == 3 || i == 5 || i == 7 || i == 9 || i == 11 || i == 13) && mas[i] == "")
                        {
                        }
                        else
                        {
                            if (mas[i] == "r")
                            { }
                            else
                            {
                                if (mas[i] != "" && mas[i] != null)
                                {
                                    k = mas[i].ToLower().Contains("практика");
                                }
                                if (k)
                                {
                                    j += $"{mas[i]}\n";
                                    break;

                                }
                                if (mas[i] == null || mas[i] == "")
                                {
                                    //SendMessage($"{i} пара - нет пары", userID);
                                    j += $"{i / 2} пара - нет пары \n[{zvon[(i / 2) - 1]}]\n\n";
                                }
                                else
                                {
                                    //SendMessage($"{i} пара - {mas[i]}", userID);
                                    j += $"{i / 2} пара - {mas[i]} \n[{zvon[(i / 2) - 1]}]\n\n";
                                    
                                }
                            }
                        }


                    }
                }
                SendMessage(j, userID);
            }
            catch { }
        }
        public static void SendMessage(string message, long? userID)
        {
            try
            {


                Random rnd = new Random();
                api.Messages.Send(new MessagesSendParams
                {
                    RandomId = rnd.Next(),
                    UserId = userID,
                    Message = message
                });
            }
            catch { }

        }
    }
}//ссылка на ексель файл http://serp-koll.ru/images/ep/k1/rasp1.xlsx