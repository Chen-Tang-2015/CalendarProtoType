using System.Collections.Generic;
using System;

namespace CalendarWebAppPrototype.Models
{
    public class MyEvent
    {
        public string Subject { get; set; }
        public string StartTimeZone { get; set; }
        public string EndTimeZone { get; set; }
        public DateTimeOffset Start { get; set; }
        public DateTimeOffset End { get; set; }
        public BodyClass Body { get; set; }
        public List<Attendee> Attendees { get; set; }
        
        public string GetAttendeeList()
        {
            string result = "";
            foreach (var attendee in Attendees)
            {
                result += (result.Length > 0 ? ";" : "") + attendee.EmailAddress.Name;
            }
            return result;            
        }
        public OrganizerClass Organizer { get; set; }
        public LocationClass Location { get; set; }
        public RecurrenceClass Recurrence { get; set; }
        public class BodyClass
        {
            public string Content { get; set; }
            public string ContentType { get; set; }
        }
        public class Attendee
        {
            public Email EmailAddress { get; set; }
        }
        public class Email
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }
        public class OrganizerClass
        {
            public Email EmailAddress { get; set; }
        }
        public class LocationClass
        {
            public string DisplayName { get; set; } 
        }
        public class RecurrenceClass
        {
            public PatternClass Pattern;
            public RangeClass Range;

        }
        public class PatternClass
        {
            public string Type;
            public int Interval;
            public int Month;
            public string FirstDayOfWeek;
            public int DayOfMonth;
            public string[] DaysOfWeek;
        }

        public class RangeClass
        {
            public string Type;
            public string StartDate;
            public string EndDate;
            public int NumberOfOccurrences;
        }
    }
}