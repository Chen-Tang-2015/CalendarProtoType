namespace CalendarWebAppPrototype.Models
{
    public class InputEvent
    {
        public string Attendee_Email { get; set; }
        public string Attendee_Email_Name { get; set; }
        public string Content { get; set; }
        public string Content_Type { get; set; }
        public string Start { get; set; }
        public string Start_Time_Zone { get; set; }
        public string End { get; set; }
        public string End_Time_Zone { get; set; }
        public string Location { get; set; }
        public string Organizer_Email { get; set; }
        public string Organizer_Email_Name { get; set; }
        public string Subject { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public int Number_Of_Occurences { get; set; }
        public string Type { get; set; }
        public int Interval { get; set;}
        public int Month { get; set; }
        public string Index { get; set; }
        public int First_Day_Of_Week { get; set; }
        public int Day_Of_Month { get; set; }    
        public string Days_of_Week { get; set; } 
        public string Range_Type { get; set; }
    }
}