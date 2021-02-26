using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows;

namespace Connect
{
        public class Utility
    {

        public static ISldWorks SwApp { get; private set; }

        public static ISldWorks ConnectToSolidWorks()
        {
            if (SwApp != null)
            {
                return SwApp;
            }
            else
            {
                Debug.Print("connect to solidworks on " + DateTime.Now);
                try
                {
                    SwApp = (ISldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
                    //SwApp.SendMsgToUser("From 2018 -1");
                }
                catch (COMException)
                {
                    try
                    {
                        SwApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application.23"));
                    }
                    catch (COMException)
                    {
                        try
                        {
                            SwApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application.26"));
                        }
                        catch (COMException)
                        {
                            MessageBox.Show("Could not connect to SolidWorks.", "SolidWorks");
                            SwApp = null;
                        }
                    }
                }
                return SwApp;
            }
        }
    }

}