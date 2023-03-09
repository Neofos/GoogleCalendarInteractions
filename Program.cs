using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CreateEvent
{
    class Program
    {
        static Program()
        {
            string? pathToCredentials = Environment.GetEnvironmentVariable("KEY_PATH");

            string[] scopes = new string[] { "https://www.googleapis.com/auth/calendar" };

            UserCredential credential;
            using (var stream = new FileStream(pathToCredentials, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    System.Threading.CancellationToken.None).Result;
            }

            string ApplicationName = "КОЛЕНДАРЕК";

            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            _service = service;
        }

        static CalendarService _service;

        const string CALENDAR_ID = "*СЮДА АЙДИ КАЛЕНДАРЯ*";

        static string clientEmail = "*СЮДА ЕМЕЙЛ*";

        static void Main(string[] args)
        {
            while (true)
            {
                if (Console.ReadLine() == "/график")
                    ShowFreeDays();
                else
                    Console.Clear();
            }
        }

        static void ShowFreeDays()
        {
            DateTime now = DateTime.Now;
            EventsResource.ListRequest request = _service.Events.List(CALENDAR_ID);
            request.TimeMin = now;
            request.TimeMax = now.AddDays(14);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            Events events = request.Execute();

            if (events.Items != null && events.Items.Count > 0)
            {
                int index = 0;
                IEnumerable<Event>? freeEvents = events.Items.Where(
                        e => e.Attendees == null || e.Attendees.Count == 0);

                if (freeEvents.Any())
                {
                    Console.WriteLine("\nНезанятые дни в ближайшие две недели:\n");

                    foreach (var eventItem in freeEvents)
                    {
                        string? whenStart = eventItem.Start.DateTime.ToString(),
                            whenEnd = eventItem.End.DateTime.Value.TimeOfDay.ToString();

                        if (string.IsNullOrEmpty(whenStart))
                            whenStart = eventItem.Start.Date;
                        if (string.IsNullOrEmpty(whenEnd))
                            whenEnd = eventItem.End.Date;

                        Console.WriteLine($"{index + 1}. ({eventItem.Start.DateTime.Value.DayOfWeek} " +
                            $"{whenStart} - {whenEnd})");

                        index++;
                    }

                    Console.Write("\nВведите номер позиции, на которую хотите сделать запись,\n" +
                        "или что-то другое, если хотите выйти: ");

                    string? userInput = Console.ReadLine();

                    if (int.TryParse(userInput, out int number) && number > 0 && number <= index + 1)
                    {
                        Console.WriteLine();
                        MakeAnAppointment(freeEvents.ElementAt(number - 1));
                    }
                    else
                    {
                        Console.WriteLine("Ну бб тогда.\n");
                    }
                }
                else
                    Console.WriteLine("\nА занято всё, лох\n");
            }
        }

        static void MakeAnAppointment(Event eventItem)
        {
            Event newEvent = new Event()
            {
                Summary = eventItem.Summary,
                Description = eventItem.Description,
                Location = eventItem.Location,
                Start = eventItem.Start,
                End = eventItem.End,
                GuestsCanModify = eventItem.GuestsCanModify,
                GuestsCanInviteOthers = eventItem.GuestsCanInviteOthers
            };

            string newCalendarId = "*СЮДА АЙДИ НОВОГО КАЛЕНДАРЯ*";

            newEvent = _service.Events.Insert(newEvent, newCalendarId).Execute();

            if (newEvent != null)
                _service.Events.Delete(CALENDAR_ID, eventItem.Id).Execute();

            if (newEvent.Attendees == null)
            {
                newEvent.Attendees = new List<EventAttendee>();
            }

            EventAttendee newAttendee = new EventAttendee();
            newAttendee.Email = clientEmail;

            newEvent.Attendees.Add(newAttendee);

            newEvent.ColorId = "11";

            var updateRequest = _service.Events.Update(newEvent, newCalendarId, newEvent.Id);
            updateRequest.SendUpdates = EventsResource.UpdateRequest.SendUpdatesEnum.All;
            Event updatedEvent = updateRequest.Execute();

            DateTime? eventStart = updatedEvent.Start.DateTime;
            Console.WriteLine($"Записал тебя на {eventStart.Value.DayOfWeek} {eventStart.Value.ToString("d")} " +
                $"в {eventStart.Value.TimeOfDay}\n");
        }

        static Event ExampleEventCreate()
        {
            EventAttendee[] attendees = new EventAttendee[]
            {
                new EventAttendee() { Email = "*емейл*" },
                new EventAttendee() { Email = "*емейл*" }
            };

            Event newEvent = new Event()
            {
                Summary = "ТЕСТЮ ИВЕНТОЧКИ",
                Location = "800 Howard St., San Francisco, CA 94103",
                Description = "A chance to be a [[BIG SHOT]]",
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2023-02-27T19:20:00"),
                    TimeZone = "Europe/Moscow",
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2023-02-27T20:00:00"),
                    TimeZone = "Europe/Moscow",
                },
                Visibility = "private",
                GuestsCanModify = false,
                Attendees = attendees
            };

            newEvent.Reminders = new Event.RemindersData();
            newEvent.Reminders.UseDefault = false;
            newEvent.Reminders.Overrides = new System.Collections.Generic.List<EventReminder>
            {
                new EventReminder() { Method = "popup", Minutes = 2 },
                new EventReminder() { Method = "email", Minutes = 2 }
            };

            // Insert event into calendar
            EventsResource.InsertRequest request = _service.Events.Insert(newEvent, CALENDAR_ID);
            request.SendNotifications = true;
            return request.Execute();
        }

        static Event EventUpdate(Event createdEvent)
        {
            createdEvent.Start = new EventDateTime()
            {
                DateTime = new DateTime(2023, 3, 4, 10, 0, 0),
                TimeZone = "Europe/Moscow"
            };
            createdEvent.End = new EventDateTime()
            {
                DateTime = new DateTime(2023, 3, 4, 16, 0, 0),
                TimeZone = "Europe/Moscow"
            };

            EventsResource.UpdateRequest requestOnceAgain = _service.Events.Update(createdEvent,
                CALENDAR_ID, // calendarId
                createdEvent.Id); // eventId
            return requestOnceAgain.Execute();
        }

        static void StartWatchingEvents()
        {
            var request = _service.Events.Watch(new Channel()
            {
                Id = Guid.NewGuid().ToString(), // Уникальный идентификатор канала
                Type = "web_hook", // Тип канала
                Address = @"" // URL для получения уведомлений
            }, CALENDAR_ID); // ID календаря

            // Выполнение запроса и получение ответа
            var response = request.Execute();

            // Вывод информации о созданном канале
            Console.WriteLine("Channel ID: {0}", response.Id);
            Console.WriteLine("Resource ID: {0}", response.ResourceId);
            Console.WriteLine("Expiration: {0}", response.Expiration);
        }
    }
}