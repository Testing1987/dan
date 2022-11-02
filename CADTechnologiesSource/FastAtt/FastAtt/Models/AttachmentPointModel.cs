using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastAtt.Models
{
    public class AttachmentPointModel
    {
        public AttachmentPoint TopLeft { get; set; } = AttachmentPoint.TopLeft;
        public AttachmentPoint TopCenter { get; set; } = AttachmentPoint.TopCenter;
        public AttachmentPoint TopRight { get; set; } = AttachmentPoint.TopRight;
        public AttachmentPoint MiddleLeft { get; set; } = AttachmentPoint.MiddleLeft;
        public AttachmentPoint MiddleCenter { get; set; } = AttachmentPoint.MiddleCenter;
        public AttachmentPoint MiddleRight { get; set; } = AttachmentPoint.MiddleRight;
        public AttachmentPoint BottomLeft { get; set; } = AttachmentPoint.BottomLeft;
        public AttachmentPoint BottomCenter { get; set; } = AttachmentPoint.BottomCenter;
        public AttachmentPoint BottomRight { get; set; } = AttachmentPoint.BottomRight;
    }
}
