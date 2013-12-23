using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AmberMeClient
{
    public class StreamManager
    {
        public Stream GetStreamByName(string fileName) {
            var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            return stream;
        }

        public Stream GetWriteStream(string fileName) {
            var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            return stream;
        }

    }
}
