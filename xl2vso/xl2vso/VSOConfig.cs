﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xl2vso
{
    class VSOConfig
    {
        public static void Initialize()
        {
            throw new NotImplementedException("initialize the data");
        }

        public static string  userName { get; private set; }
        public static string password { get; private set; }
        public static string accountName { get; private set; }
        public static string projectName { get; private set; }
    }
}
