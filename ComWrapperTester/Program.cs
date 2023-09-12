using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using camLineV2.Com;

namespace ComWrapperTester
{
    internal class Program
    {
        static void Main(string[] args)
        {


            camLineV2ComWrapper test = new camLineV2ComWrapper();
            
            string error = null;
            string json = null;
            string totalCount = null;
            string totalFail = null;
            string totalPass = null;
            string totalScrap = null;
            string parameter = null;

            int test2 = test.Init("TEST", ref json, ref error);
            string status = null;
            string  units=  "100309300107102301600330" ;


            // var statustest = null;

            string status2 = null;

            json = json.Replace("RBLDD2MW_TEST", "RBT4033N_VERP");


            //int  test3 = test.EQP_CheckUnit("",units , ref status2 , ref error);

            


          // var test4 = test.EQP_CheckUnit_Single("100309300107102301600330", "ICAS_SEMI1", ref status,ref error);

          //var test5 = test.EQP_GetOwnProcessCounters("", units, ref error, json, ref  totalCount, ref  totalFail,ref totalFail, ref  totalScrap);

          var test6 = test.EQP_GetParameter("18414301018243323041", "MRA2_SN", json, "p_passcount", ref parameter, ref error);

          int test4 = test.EQP_CheckStationStatus(json, ref error);

            Console.ReadKey();

        }
    }
}
