// <copyright file="BaseNotification.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;
    using Microsoft.OData.Edm;

    /// <summary>
    /// Base notification model class.
    /// </summary>
    public class BaseNotification
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Secondary button Title value.
        /// </summary>
        public string ButtonTitle2 { get; set; }

        /// <summary>
        /// Gets or sets the Secondary button Link value.
        /// </summary>
        public string ButtonLink2 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether scheduled or not.
        /// </summary>
        public bool IsScheduled { get; set; }

        /// <summary>
        /// Gets or sets the Schedule DateTime value.
        /// </summary>
        public DateTime ScheduleDate { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether recurrence message or not.
        /// </summary>
        public bool IsRecurrence { get; set; }

        /// <summary>
        /// Gets or sets the Repeats value (EveryWeekday/Daily/Weekly/Monthly/Yearly/Custom).
        /// </summary>
        public string Repeats { get; set; }

        /// <summary>
        /// Gets or sets the Repeat for value.
        /// </summary>
        public int RepeatFor { get; set; }

        /// <summary>
        /// Gets or sets the Repeat frequency value (Day/Week/Month).
        /// </summary>
        public string RepeatFrequency { get; set; }

        /// <summary>
        /// Gets or sets the Week Selection value (0/1/2/3/4/5/6).
        /// </summary>
        public string WeekSelection { get; set; }

        /// <summary>
        /// Gets or sets the repeat StartDate DateTime value.
        /// </summary>
        public DateTime RepeatStartDate { get; set; }

        /// <summary>
        /// Gets or sets the repeat EndDate DateTime value.
        /// </summary>
        public DateTime RepeatEndDate { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }
    }
}
