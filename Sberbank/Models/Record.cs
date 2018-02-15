using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace Sberbank.Models
{
    // Add profile data for application users by adding properties to the ApplicationUser class
    public class Record
    {
        public int RecordId { get; set; }
        [DataType(DataType.Date)]
        public DateTime date { get; set; }
        public double earnings { get; set; }
        [Range(0, 99999.99)]
        public float currency { get; set; }
        [Range(0, 9999.99)]
        public float index { get; set; }
    }
}
