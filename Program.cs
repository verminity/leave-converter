using System.Globalization;
using System.Text.RegularExpressions;
using ExcelDataReader;

internal class Program
{
    const string DateFormat = "dd/MM/yyyy";
    const string FullDayPattern = """^\d{2}/\d{2}/\d{4}$""";
    const string HalfDayPattern = """^\d{2}/\d{2}/\d{4} (am|pm)$""";
    const string DayRangePattern = """^\d{2}/\d{2}/\d{4} - \d{2}/\d{2}/\d{4}$""";

    private static void Main(string[] args)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(args[0], FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var allLeaveDates = GetLeaveDates(reader)
            .SelectMany(LeaveStringToLeaveDates)
            .GroupBy(x => x.WeekOfYear)
            .OrderBy(x => x.Key)
            .ToList();

        var firstWeek = allLeaveDates.First().Key;
        var endWeek = allLeaveDates.Last().Key;

        var weeklyLeave = firstWeek.Until(endWeek)
        .Select(x =>
        {
            var matchingLeave = allLeaveDates.FirstOrDefault(w => w.Key == x);
            if (matchingLeave is null) return "";
            return matchingLeave.Sum(w => w.AmountOfLeave).ToString();
        }).ToList();

        Console.WriteLine($"Copy the following line and paste it into week {firstWeek.Week} of {firstWeek.Year} and then \"Split text to columns\" to populate the cells");

        Console.WriteLine(string.Join(',', weeklyLeave));
    }

    static IEnumerable<string> GetLeaveDates(IExcelDataReader reader)
    {
        do
        {
            while (reader.Read())
            {
                yield return reader.GetString(0);
            }

        } while (reader.NextResult());
    }

    static IEnumerable<LeaveDate> LeaveStringToLeaveDates(string leaveString)
    {
        if (Regex.IsMatch(leaveString, FullDayPattern))
        {
            yield return new LeaveDate(DateOnly.ParseExact(leaveString, DateFormat), 1);
            yield break;
        }

        if (Regex.IsMatch(leaveString, HalfDayPattern))
        {
            yield return new LeaveDate(DateOnly.ParseExact(leaveString.Substring(0, 10), DateFormat), 0.5);
            yield break;
        }

        if (Regex.IsMatch(leaveString, DayRangePattern))
        {
            var currentDate = DateOnly.ParseExact(leaveString.Substring(0, 10), DateFormat);
            var endDate = DateOnly.ParseExact(leaveString.Substring(13, 10), DateFormat);

            do
            {
                if (!(currentDate.DayOfWeek == DayOfWeek.Saturday || currentDate.DayOfWeek == DayOfWeek.Sunday))
                {
                    yield return new LeaveDate(currentDate, 1);
                }
                currentDate = currentDate.AddDays(1);
            } while (currentDate <= endDate);
            yield break;
        }

        Console.WriteLine($"Failed to parse cell containing '{leaveString}'");
    }
}

record WeekOfYear(int Week, int Year) : IComparable<WeekOfYear>
{
    public int CompareTo(WeekOfYear? other)
    {
        if (other == null) { return 0; }
        if (other.Year == Year) { return Week.CompareTo(other.Week); }
        return Year.CompareTo(other.Year);
    }

    public IEnumerable<WeekOfYear> Until(WeekOfYear endWeek)
    {
        var currentWeek = this;
        do
        {
            yield return currentWeek;
            currentWeek = currentWeek.Next();
        } while (currentWeek != endWeek);
        yield return endWeek;
    }

    WeekOfYear Next()
    {
        var maxWeeks = ISOWeek.GetWeeksInYear(Year);

        if (Week == maxWeeks)
        {
            return new WeekOfYear(1, Year + 1);
        }
        return new WeekOfYear(Week + 1, Year);
    }
}

record LeaveDate
{
    public WeekOfYear WeekOfYear { get; private init; }
    public DateOnly Date { get; }
    public double AmountOfLeave { get; }

    public LeaveDate(DateOnly date, double amountOfLeave)
    {
        Date = date;
        AmountOfLeave = amountOfLeave;

        var week = ISOWeek.GetWeekOfYear(date.ToDateTime(TimeOnly.MinValue));
        var year = ISOWeek.GetYear(date.ToDateTime(TimeOnly.MinValue));

        WeekOfYear = new WeekOfYear(week, year);
    }
}


