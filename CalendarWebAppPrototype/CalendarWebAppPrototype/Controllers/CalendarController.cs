using Microsoft.IdentityModel.Clients.ActiveDirectory;
using CalendarWebAppPrototype.Utils;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Mvc;
using CalendarWebAppPrototype.Models;
using System;
using System.Dynamic;
using System.Web;


namespace CalendarWebAppPrototype.Controllers
{
    public class CalendarController : Controller
    {
        static int pageSize = 10;
        static int skipSize = 10;
       
        // GET: Index
        [Authorize]
        public async Task<ActionResult> Index()
        {
            return View();
        }

        // GET: Event
        [Authorize]
        public async Task<ActionResult> Events(string nextLink = null)
        {
            List<MyEvent> myEvents = new List<MyEvent>();
            try
            {
                if (string.IsNullOrEmpty(nextLink))
                {
                    myEvents = await GetEventsWithinTimeFrame(DateTime.Parse(Request.QueryString["start"]), DateTime.Parse(Request.QueryString["end"]));
                }
                else
                {
                    myEvents = await GetEventsWithinTimeFrame(new DateTime(), new DateTime(), nextLink);
                } 
            }
            catch (Exception)
            {
                return RedirectToAction("Index");
            }
            return View(myEvents);
        }

        private async Task<List<MyEvent>> GetEventsWithinTimeFrame(DateTime start, DateTime end, string uri = null)
        {
            List<MyEvent> myEvents = new List<MyEvent>();
            try
            {
                if (string.IsNullOrEmpty(uri))
                {
                    uri = Util.baseUri + "calendarview?startDateTime=" + start.ToUniversalTime().ToString("yyyy-MM-dd\"T\"HH:mm:ssZ") + "&endDateTime=" + end.ToUniversalTime().ToString("yyyy-MM-dd\"T\"HH:mm:ssZ") + "&$Select=Subject,Attendees,Start,End,StartTimeZone,EndTimeZone,Body,Organizer,Location&$orderby=start&$top=" + pageSize +"&$skip=" + 0;
                }
                myEvents = await Util.GetItemsAsync<List<MyEvent>>(uri);
                //string nextLink = Response["odata.nextLink"];
                //var outputEvent = await Util.GetItemAsync<IDictionary<string, object>>(Uri);
                //object obj1, obj2;
                //outputEvent.TryGetValue("@odata.nextLink", out obj1);
                //string nextLink = (string) obj1;
                //if (!string.IsNullOrEmpty(nextLink))
                //{
                //    ViewData["nextLink"] = nextLink;
                //}
                //outputEvent.TryGetValue("value", out obj2);
                //var events = (JArray) obj2;
                //myEvents = events.ToObject<List<MyEvent>>();
                if (myEvents.Count == pageSize)
                {
                    Uri uri2 = new Uri(uri);
                    int originalSkip = int.Parse(HttpUtility.ParseQueryString(uri2.Query).Get("$skip"));
                    string startDateTime = HttpUtility.ParseQueryString(uri2.Query).Get("startDateTime");
                    string endDateTime = HttpUtility.ParseQueryString(uri2.Query).Get("endDateTime");
                    int nextSkip = originalSkip + skipSize;
                    ViewData["nextLink"] = Util.baseUri + "calendarview?startDateTime=" + startDateTime + "&endDateTime=" + endDateTime+ "&$Select=Subject,Attendees,Start,End,StartTimeZone,EndTimeZone,Body,Organizer,Location&$orderby=start&$top=" + pageSize + "&$skip=" + nextSkip;
                }
            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    ViewBag.ErrorMessage = "AuthorizationRequired";
                }
            }
            return myEvents;
        }

        [Authorize]
        public ActionResult CreateMeeting()
        {
            return View();
        }

        [Authorize][HttpPost]
        public async Task<ActionResult> CreateMeeting(InputEvent InputEvent)
        {
            await Util.PostItemAsync<ExpandoObject>("events", FillEvent(InputEvent));
            return RedirectToAction("Index");
        }

        [Authorize]
        public async Task<ActionResult> ListInvitation()
        {
            List<EventMessage> eventMessages = await GetEventMessages();
            return View(eventMessages);
        }
                
        private async Task<List<EventMessage>> GetEventMessages()
        {
            List<EventMessage> eventMessages = new List<EventMessage>();
            List<EventMessage> messages = new List<EventMessage>();
            List<EventMessage> tmp = new List<EventMessage>();
            try
            {
                string uri = null;
                int top = 50, skipSize = 50;
                int skip = 0;
                do
                {
                    uri = Util.baseUri + "folders/Inbox/messages?&$orderby=DateTimeReceived%20desc&$top=" + top + "&$skip=" + skip;
                    tmp = await Util.GetItemsAsync<List<EventMessage>>(uri);
                    messages.AddRange(tmp);                  
                    skip += skipSize;
                } while (skip <= 500);
                //string nextLink = Response["odata.nextLink"];
                //var outputEvent = await Util.GetItemAsync<IDictionary<string, object>>(Uri);
                //object obj1, obj2;
                //outputEvent.TryGetValue("@odata.nextLink", out obj1);
                //string nextLink = (string) obj1;
                //if (!string.IsNullOrEmpty(nextLink))
                //{
                //    ViewData["nextLink"] = nextLink;
                //}
                //outputEvent.TryGetValue("value", out obj2);
                //var events = (JArray) obj2;
                //myEvents = events.ToObject<List<MyEvent>>();   
                eventMessages  = (messages.FindAll(
                  delegate (EventMessage m)
                  {
                      return "#Microsoft.OutlookServices.EventMessage".Equals(m.Type);

                     //  return m.Sender.EmailAddress.Address == "v-chenta@microsoft.com";
                     // return true;
                  }
                  ));

            }
            catch (AdalException exception)
            {
                //handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    ViewBag.ErrorMessage = "AuthorizationRequired";
                }
            }
            return eventMessages;
        }


        private ExpandoObject FillEvent(InputEvent InputEvent)
        {
            dynamic e = new ExpandoObject();
            e.Start = new DateTimeOffset(DateTime.Parse(InputEvent.Start));
            e.End = new DateTimeOffset(DateTime.Parse(InputEvent.End));
            e.StartTimeZone = InputEvent.Start_Time_Zone;
            e.EndTimeZone = InputEvent.End_Time_Zone;
            e.Subject = InputEvent.Subject;
            e.Organizer = new ExpandoObject();
            e.Organizer.EmailAddress = new ExpandoObject();
            e.Organizer.EmailAddress.Address = InputEvent.Organizer_Email;
            e.Organizer.EmailAddress.Name = InputEvent.Organizer_Email_Name;
            e.Attendees = new List<ExpandoObject>();
            e.Attendees.Add(e.Organizer);
            string [] attendeesEmail = InputEvent.Attendee_Email.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string [] attendeeEmailName = InputEvent.Attendee_Email_Name.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < attendeesEmail.Length; i++)
            {
                dynamic attendee = new ExpandoObject();
                attendee.EmailAddress = new ExpandoObject();
                attendee.EmailAddress.Address = attendeesEmail[i];
                attendee.EmailAddress.Name = attendeeEmailName[i];
                e.Attendees.Add(attendee);
            }
            e.Body = new ExpandoObject();
            e.Body.Content = InputEvent.Content;
            e.Body.ContentType = InputEvent.Content_Type;
            e.Location = new ExpandoObject();
            e.Location.DisplayName = InputEvent.Location;
            if (!string.IsNullOrEmpty(InputEvent.Type))
            {
                e.Recurrence = new ExpandoObject();
                e.Recurrence.Pattern = new ExpandoObject();
                e.Recurrence.Pattern.Type = InputEvent.Type;
                e.Recurrence.Pattern.Interval = InputEvent.Interval;
                if (InputEvent.Type.Equals("Daily"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Month = 0;
                    e.Recurrence.Pattern.Index = "First";
                    e.Recurrence.Pattern.DayOfMonth = 0;
                }
                else if (InputEvent.Type.Equals("Weekly"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Month = 0;
                    e.Recurrence.Pattern.DayOfMonth = 0;
                    e.Recurrence.Pattern.Index = "First";
                    e.Recurrence.Pattern.DaysOfWeek = InputEvent.Days_of_Week.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                }
                else if (InputEvent.Type.Equals("RelativeMonthly"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Month = 0;
                    e.Recurrence.Pattern.DayOfMonth = 0;
                    e.Recurrence.Pattern.Index = InputEvent.Index;
                    e.Recurrence.Pattern.DaysOfWeek = InputEvent.Days_of_Week.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                }
                else if (InputEvent.Type.Equals("AbsoluteMonthly"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Index = "First";
                    e.Recurrence.Pattern.Month = 0;
                    e.Recurrence.Pattern.DayOfMonth = InputEvent.Day_Of_Month;
                }
                else if (InputEvent.Type.Equals("AbsoluteYearly"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Index = "First";
                    e.Recurrence.Pattern.Month = InputEvent.Month;
                    e.Recurrence.Pattern.DayOfMonth = InputEvent.Day_Of_Month;
                }
                else if (InputEvent.Type.Equals("RelativeYearly"))
                {
                    e.Recurrence.Pattern.FirstDayOfWeek = "Sunday";
                    e.Recurrence.Pattern.Month = InputEvent.Month;
                    e.Recurrence.Pattern.DayOfMonth = 0;
                    e.Recurrence.Pattern.Index = InputEvent.Index;
                    e.Recurrence.Pattern.DaysOfWeek = InputEvent.Days_of_Week.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                }

                e.Recurrence.Range = new ExpandoObject();
                e.Recurrence.Range.StartDate = InputEvent.StartDate;
                e.Recurrence.Range.Type = InputEvent.Range_Type;
                if (InputEvent.Range_Type.Equals("NoEnd"))
                {
                    e.Recurrence.Range.EndDate = "0001-01-01T00:00:00Z";
                    e.Recurrence.Range.NumberOfOccurrences = 0;
                }
                else if (InputEvent.Range_Type.Equals("Numbered"))
                {
                    e.Recurrence.Range.EndDate = "0001-01-01T00:00:00Z";
                    e.Recurrence.Range.NumberOfOccurrences = InputEvent.Number_Of_Occurences;
                }
                else if (InputEvent.Range_Type.Equals("EndDate"))
                {
                    e.Recurrence.Range.EndDate = InputEvent.EndDate;
                    e.Recurrence.Range.NumberOfOccurrences = 0;
                }               
            }
            return e;
        }    
    }  
}
