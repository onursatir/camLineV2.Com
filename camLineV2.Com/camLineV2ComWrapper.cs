using camLineV2.iServices.Equipment;
using camLineV2.StatusCode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.XPath;
using NLog.Targets;

namespace camLineV2.Com
{

    [ComVisible(true)]
    [Guid("24702C3E-3D7C-4369-ABEC-8AAC551E10EB")]
    [ProgId("Continental.camLineV2ComWrapper")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IcamLineV2ComWrapper))]
    public class camLineV2ComWrapper : IcamLineV2ComWrapper

    {
        private string[,] unitStatus_out;
        private string[,] unitParameters_out;
        private string settingsAsJSON = "";

        public int Init(string process, ref string settingsAsJSON, ref string errorDescription)
        {
            errorDescription = null;
            int result = -100;

            var instance = DIContainer.GetService<IcamLineSimpleMES>();
            if (instance == null)
            {
                result = Status.TEW_V2SimpleInstanceNull(out errorDescription);
                return result;
            }
            else
            {
                result = instance.Init(out errorDescription, process, out settingsAsJSON);
                return result;
            }
        }


        //MopsId: 

        public int EQP_CheckUnit(string mopsUnitId,  string singleUnitId, ref string UnitStatus, ref string errorDescription)
        {
            UnitStatus = "";
            errorDescription = "";
            int result = -100;
            string[] units ={""};


            if (!settingsAsJSON.Contains("\"UnitType\": \"mops\","))
            {
                units[0] = singleUnitId;
            }

            var instance = DIContainer.GetService<IcamLineSimpleMES>();
            if (instance == null)
            {
                result = Status.TEW_V2SimpleInstanceNull(out errorDescription);

            }
            else
            {
                result = instance.EQP_CheckUnit_01Simple(out errorDescription, settingsAsJSON, mopsUnitId, units, out unitStatus_out);

                int row = unitStatus_out.GetLength(0);
                if (result == 0)

                {
                    for (int i = 1; i < row; i++)
                    {
                        UnitStatus = unitStatus_out[row - 1, 2].ToString() + ",";
                        errorDescription = unitStatus_out[row - 1, 7].ToString() + ",";
                    }

                    UnitStatus = UnitStatus.Substring(0, UnitStatus.Length - 1);
                    errorDescription = errorDescription.Substring(0, errorDescription.Length - 1);
              
                }

            }
            return result;
        }



        // Rictig ist es mit Array Rückgabe, Hier besteht ein Klärungsbedarf,daher 
        //public int EQP_CheckUnit(string Id, string IdType, ref string[] UnitStatus)


        //{
        //    UnitStatus = null;
        //    var instance = DIContainer.GetService<IcamLineSimpleMES>();
        //    var dict = new Dictionary<int, string>();
        //    var errorDescription = "";
        //    string mopsUnitId = null;
        //    string[] units = null;
        //    int result = -100;

        //    if (!settingsAsJSON.Contains("\"UnitType\": \"mops\","))
        //    {
        //        units = new[] { Id + "," + IdType };
        //    }

        //    else
        //    {
        //        mopsUnitId = Id + IdType;
        //    }
        //    if (instance == null)
        //    {
        //        result = Status.TEW_V2SimpleInstanceNull(out errorDescription);


        //    }
        //    else
        //    {
        //        result = instance.EQP_CheckUnit_01Simple(out errorDescription, settingsAsJSON, mopsUnitId, units, out unitStatus_out);

        //        int row = unitStatus_out.GetLength(0);
        //        if (result == 0)


        //        {

        //            for (int i = 1; i < row; i++)
        //            {
        //                UnitStatus = new[] { unitStatus_out[row - 1, 2].ToString() };
        //            }




        //        }



        //    }
        //    return result;
        //}



        public int EQP_CheckStationStatus(string settingsAsJSON, ref string errorDescription)
        {
            int result = -100;
            try
            {
                

                var instance = DIContainer.GetService<IcamLineSimpleMES>();
                if (instance == null)
                {
                    result=  Status.TEW_V2SimpleInstanceNull(out errorDescription);

                }
                else
                {
                    result= instance.EQP_CheckStationStatus_01Simple(out errorDescription, settingsAsJSON);

                }
            }
            catch (Exception ex)
            {
                result = Status.TEW_GeneralParsingMappingError(out errorDescription, ex.Message);
            }
            return result;

        }

      
        public int EQP_GetOwnProcessCounters(string mopsUnitId, string singleUnitId, string settingsAsJSON, ref string totalCount, ref string totalFail, ref string totalPass, ref string totalScrap, ref string errorDescription )
        {

            string[] units = { "" };
            int result = -100;


            if (!settingsAsJSON.Contains("\"UnitType\": \"mops\","))
            {
                units[0] = singleUnitId;
            }

            var instance = DIContainer.GetService<IcamLineSimpleMES>();
            if (instance == null)
            {
                result = Status.TEW_V2SimpleInstanceNull(out errorDescription);

            }
            else
            {

                 result = instance.EQP_GetOwnProcessCounters_01Simple(out errorDescription, settingsAsJSON, mopsUnitId, units, out string[,] unitsCounters_out);

                int row = unitsCounters_out.GetLength(0);
                if (result == 0)
                {
                    for (int i = 1; i < row; i++)
                    {
                        totalCount =unitsCounters_out[row - 1, 1].ToString() + ",";
                        totalFail = unitsCounters_out[row - 1, 2].ToString() + ",";
                        totalPass = unitsCounters_out[row - 1, 3].ToString() + ",";
                        totalScrap = unitsCounters_out[row - 1, 4].ToString() + ",";
                    }

                    totalCount = totalCount.Substring(0, totalCount.Length - 1);
                    totalFail = totalFail.Substring(0, totalFail.Length - 1);
                    totalPass = totalPass.Substring(0, totalPass.Length - 1);
                    totalScrap = totalScrap.Substring(0, totalScrap.Length - 1);
                }

               
            }

            return result;


        }


        public int EQP_GetParameter(string Id, string IdType,string settingsAsJSON, string parameter, ref string parameterInhalt, ref string errorDescription)
        {

            errorDescription = "";
            int result = -100;
            parameterInhalt = null;

            var instance = DIContainer.GetService<IcamLineSimpleMES>();
            if (instance == null)
            {
                result = Status.TEW_V2SimpleInstanceNull(out errorDescription);

            }
            else
            {
                result = instance.EQP_GetParameters_01Simple(out errorDescription, settingsAsJSON, Id, IdType, parameter, out unitParameters_out);

                parameterInhalt = unitParameters_out[1, 3];

            }

            return result;

        }
    }
}