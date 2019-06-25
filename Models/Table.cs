using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NPOI_API.Models
{
    public class Table
    {
        public int id { get; set; }
        public string SectionDescription { get; set; }
        public double PPM { get; set; }
        public double MetCriteria { get; set; }
    }
}