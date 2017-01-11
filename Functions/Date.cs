using System;
using System.Collections.Generic;

namespace CalcEngine.Functions
{
    internal static class Date
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("DAYS", 2, Days);
            ce.RegisterFunction("NOW", 0, Now);
            ce.RegisterFunction("YEARS", 2, Years);
            ce.RegisterFunction("NUMBEROFLEAPDAYS", 2, NumberOfLeapDays);
            ce.RegisterFunction("TOSHORTDATESTRING", 1, ToShortDateString);
        }

#if DEBUG

        public static void Test(CalcEngine ce)
        {
            ce.Test("DAYS('1/1/2000', '1/1/2001')", 366);
            //ce.Test("NOW", DateTime.Now);
            ce.Test("YEARS('1/1/2000', '1/1/2001')", 1);
            ce.Test("NUMBEROFLEAPDAYS('1/1/2000', '1/1/2005')", 2);
        }

#endif

        /// <summary>
        /// Returns the number of days between the start date and end date.
        /// </summary>
        /// <param name="p">The list of expressions/parameters (1 or more parameters)
        ///     p[0]=Start DateTime
        ///     p[n]=End DateTime
        /// </param>
        /// <returns></returns>
        private static object Days(List<Expression> p)
        {
            DateTime start = (DateTime)p[0];
            DateTime end = (DateTime)p[1];

            return end.Date.Subtract(start.Date).TotalDays;
        }

        /// <summary>
        /// Returns the current date and time.
        /// </summary>
        /// </param>
        /// <returns></returns>
        private static object Now(List<Expression> p)
        {
            return DateTime.Now;
        }

        /// <summary>
        /// Returns the number of years between the start date and end date.
        /// </summary>
        /// <param name="p">The list of expressions/parameters (1 or more parameters)
        ///     p[0]=Start DateTime
        ///     p[n]=End DateTime
        /// </param>
        /// <returns></returns>
        private static object Years(List<Expression> p)
        {
            DateTime start = (DateTime)p[0];
            DateTime end = (DateTime)p[1];

            double totalMonths = 0;
            totalMonths += (end.Year - start.Year) * 12;
            totalMonths += (end.Month - start.Month);
            totalMonths += (end.Day / (double)DateTime.DaysInMonth(end.Year, end.Month)) - (start.Day / (double)DateTime.DaysInMonth(start.Year, start.Month));
            return totalMonths / 12;
        }

        /// <summary>
        /// Returns the number of leap days between the start date and end date.
        /// </summary>
        /// <param name="p">The list of expressions/parameters (1 or more parameters)
        ///     p[0]=Start DateTime
        ///     p[n]=End DateTime
        /// </param>
        /// <returns></returns>
        private static object NumberOfLeapDays(List<Expression> p)
        {
            DateTime startDate = p[0];
            DateTime endDate = p[1];

            int numberOfLeapYears = 0;
            DateTime sDate = startDate;
            if ((sDate <= new DateTime(sDate.Year, 2, 28) && endDate > new DateTime(sDate.Year, 2, 28) && new DateTime(sDate.Year, 2, 1).AddDays(28).Month == 2) ||
            (sDate <= new DateTime(endDate.Year, 2, 28) && endDate > new DateTime(endDate.Year, 2, 28) && new DateTime(endDate.Year, 2, 1).AddDays(28).Month == 2))
            {
                numberOfLeapYears++;
            }
            while ((sDate = sDate.AddYears(1)).Year < endDate.Year)
            {
                if (DateTime.IsLeapYear(sDate.Year))
                    numberOfLeapYears++;
            }
            return numberOfLeapYears;
        }

        /// <summary>
        /// Returns short date string.
        /// </summary>
        /// <param name="p">The list of expressions/parameters (1 or more parameters)
        ///     p[0]=DateTime
        /// </param>
        /// <returns></returns>
        private static object ToShortDateString(List<Expression> p)
        {
            DateTime date = (DateTime)p[0];
            return date.ToShortDateString();
        }
    }
}