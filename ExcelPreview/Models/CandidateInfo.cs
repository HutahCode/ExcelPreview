using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ExcelPreview.Models
{
    public class CandidateInfo
    {        
        [Required]
        public double CandidateId { get; set; }

        [Required]
        public string CandidateName { get; set; }

        [Required]
        public double Exp { get; set; }

        [Required]
        public double Salary { get; set;  }

        [Required]        
        [DataType(DataType.Date), DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime DOB { get; set; }
    }
}