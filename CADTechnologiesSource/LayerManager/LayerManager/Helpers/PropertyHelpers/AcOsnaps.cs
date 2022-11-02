using System;
using System.Collections.Generic;
using System.Text;

namespace LayerManager.Helpers.PropertyHelpers
{
    public class AcOsnaps
    {
        public const int None = 0;

        public const int Endpoint = 1;

        public const int Midpoint = 2;

        public const int Center = 4;

        public const int Node = 8;

        public const int Quadrant = 16;

        public const int Intersection = 32;

        public const int Insertion = 64;

        public const int Perpendicular = 128;

        public const int Tangent = 256;

        public const int Nearest = 512;

        public const int Geometric_Center = 1024;

        public const int Apparent_Intersection = 2056;

        public const int Extension = 4096;

        public const int Parallel = 8192;

        public const int Suppress_Current = 16384;
    }
}
