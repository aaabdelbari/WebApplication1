using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class ImageFileInfo
    {
        public byte[] ImageBytes { get; set; }

        public float Width { get; set; }

        public float Height { get; set; }
    }
}