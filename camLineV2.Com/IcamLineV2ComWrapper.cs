using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace camLineV2.Com
{


    [ComVisible(true)]
    [Guid("057E7E2C-ABC9-48F4-B0E1-B7C1B990F224")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]// = default value, for late and early binding
    interface IcamLineV2ComWrapper
    {

        [DispId(1)]
        int Init(string process,  ref string json, ref string errorDescription);

        [DispId(2)]
        int EQP_CheckUnit(string mopSUnitId, string singleUnitId, ref string UnitStatus, ref string errorDescription);

        [DispId(3)]
        int EQP_CheckStationStatus( string settingsAsJSON, ref string errorDescription);

        [DispId(4)]
        int EQP_GetOwnProcessCounters(string mopsUnitId, string singleUnitId,  string settingsAsJSON, ref string totalCount, ref string totalFail, ref string totalPass, ref string totalScrap, ref string errorDescription);

        [DispId(5)]
        int EQP_GetParameter(string Id, string IdType,string settingsAsJSON,  string parameter, ref string parameterInhalt, ref string errorDescription);
        

    }


}
