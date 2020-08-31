//===============================================================================
// Microsoft FastTrack for Azure
// Azure Key Vault Samples for SharePoint Online
//===============================================================================
// Copyright © Microsoft Corporation.  All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
// FITNESS FOR A PARTICULAR PURPOSE.
//===============================================================================
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace KeyVaultSPO.Core.Models
{
    [Serializable]
    public class Post
    {
        public const string ProjectType = "Project";
        public const string TaskType = "Task";
        public const string VirtualLocation = "Virtual (Anywhere)";

        public int ID { get; set; }
        [StringLength(50)]
        [Required]
        public string Title { get; set; }
        [Required]
        public string Description { get; set; }
        [Required]
        public string Type { get; set; }
        [DisplayName("Effort in Hours")]
        [Range(0,99)]
        public int? EffortHours { get; set; }
        [DisplayName("Effort in Minutes")]
        [Range(0,59)]
        public int? EffortMinutes { get; set; }
        [DisplayName("Start Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime? StartDate { get; set; }
        [DisplayName("End Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:yyyy-MM-dd}")]
        public DateTime? EndDate { get; set; }
        [DisplayName("Expiration Date")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:yyyy-MM-dd}")]
        [Required]
        public DateTime ExpirationDate { get; set; }
        [Required]
        public string Location { get; set; }
        public List<string> Skills { get; set; }
        public string PostedBy { get; set; }
        public int PostedByID { get; set; }
        public string PostedByEmailAddress { get; set; }
        public string Status { get; set; }
        public int? AppliedID { get; set; }
        public int? SavedID { get; set; }
    }
}
