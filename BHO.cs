using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SHDocVw;
using mshtml;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices; 

namespace IEExtension
{
    /* define the IObjectWithSite interface which the BHO class will implement.
     * The IObjectWithSite interface provides simple objects with a lightweight siting mechanism (lighter than IOleObject).
     * Often, an object must communicate directly with a container site that is managing the object. 
     * Outside of IOleObject::SetClientSite, there is no generic means through which an object becomes aware of its site. 
     * The IObjectWithSite interface provides a siting mechanism. This interface should only be used when IOleObject is not already in use.
     * By using IObjectWithSite, a container can pass the IUnknown pointer of its site to the object through SetSite. 
     * Callers can also get the latest site passed to SetSite by using GetSite.
     */
    [
        ComVisible(true),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")
        // Never EVER change this UUID!! It allows this BHO to find IE and attach to it
    ]
    public interface IObjectWithSite
    {
        [PreserveSig]
        int SetSite([MarshalAs(UnmanagedType.IUnknown)]object site);
        [PreserveSig]
        int GetSite(ref Guid guid, out IntPtr ppvSite);
    }

    /* The BHO site is the COM interface used to establish a communication.
     * Define a GUID attribute for your BHO as it will be used later on during registration / installation
     */
    [
            ComVisible(true),
            Guid("2159CB25-EF9A-54C1-B43C-E30D1A4A8277"),
            ClassInterface(ClassInterfaceType.None)
    ]
    public class BHO : IObjectWithSite
    {
        private WebBrowser webBrowser;
        public const string BHO_REGISTRY_KEY_NAME =
               "Software\\Microsoft\\Windows\\" +
               "CurrentVersion\\Explorer\\Browser Helper Objects";
        /* The SetSite() method is where the BHO is initialized and where you would perform all the tasks that happen only once.
         * When you navigate to a URL with Internet Explorer, you should wait for a couple of events to make sure the required document
         * has been completely downloaded and then initialized. Only at this point can you safely access its content through the exposed
         * object model, if any. This means you need to acquire a couple of pointers. The first one is the pointer to IWebBrowser2, 
         * the interface that renders the WebBrowser object. The second pointer relates to events.
         * This module must register as an event listener with the browser in order to receive the notification of downloads
         * and document-specific events.
         */
        public int SetSite(object site)
        {
            if (site != null)
            {
                webBrowser = (WebBrowser)site;
                webBrowser.DocumentComplete +=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
            }
            else
            {
                webBrowser.DocumentComplete -=
                  new DWebBrowserEvents2_DocumentCompleteEventHandler(
                  this.OnDocumentComplete);
                webBrowser = null;
            }

            return 0;
        }

        public int GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(webBrowser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }

        public void OnDocumentComplete(object pDisp, ref object URL)
        {
            HTMLDocument document = (HTMLDocument)webBrowser.Document;

            IHTMLElement head = (IHTMLElement)((IHTMLElementCollection)
                                    document.all.tags("head")).item(null, 0);

            IHTMLScriptElement scriptObject =
                (IHTMLScriptElement)document.createElement("script");
            scriptObject.type = @"text/javascript";
            scriptObject.text = "alert('HOLA!!!');"; // <---- JAVASCRIPT INJECTION HAPPENS HERE!

            ((HTMLHeadElement)head).appendChild((IHTMLDOMNode)scriptObject);

        }

        /* The Register method simply tells IE which is the GUID of your extension so that it could be loaded.
         * The "No Explorer" value simply says that we don't want to be loaded by Windows Explorer.
         */
        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);

            if (registryKey == null)
                registryKey = Registry.LocalMachine.CreateSubKey(
                                        BHO_REGISTRY_KEY_NAME);

            string guid = type.GUID.ToString("B");
            RegistryKey ourKey = registryKey.OpenSubKey(guid);

            if (ourKey == null)
            {
                ourKey = registryKey.CreateSubKey(guid);
            }

            ourKey.SetValue("NoExplorer", 1, RegistryValueKind.DWord);

            registryKey.Close();
            ourKey.Close();
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            RegistryKey registryKey =
              Registry.LocalMachine.OpenSubKey(BHO_REGISTRY_KEY_NAME, true);
            string guid = type.GUID.ToString("B");

            if (registryKey != null)
                registryKey.DeleteSubKey(guid, false);
        }

    }
}
