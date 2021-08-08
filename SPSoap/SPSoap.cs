using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SPSoap {

    public class member_attr {
        public string name;
        public string type;
    }

    class SPSoap {
        static void Main(string[] args) {

            SPConvert cvt = new SPConvert();
            cvt.CreateSPList();
            cvt.CreateSPView();
        }
       
    }
}
