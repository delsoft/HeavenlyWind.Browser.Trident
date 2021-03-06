﻿using Microsoft.Win32;
using Sakuno.SystemInterop;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Sakuno.KanColle.Amatsukaze.Browser.Trident
{
    class TridentBrowserProvider : IBrowserProvider
    {
        TridentBrowser r_BrowserInstance;

        public IBrowser CreateBrowserInstance()
        {
            try
            {
                SetRegistry();
            }
            catch
            {
            }

            return r_BrowserInstance ?? (r_BrowserInstance = new TridentBrowser());
        }

        public void SetPort(int rpPort) => SetProxy("localhost:" + rpPort.ToString());
        static void SetProxy(string rpProxy)
        {
            var rInfo = new NativeStructs.INTERNET_PROXY_INFO()
            {
                dwAccessType = 3,
                proxy = Marshal.StringToHGlobalAnsi(rpProxy),
                proxyBypass = Marshal.StringToHGlobalAnsi("<local>")
            };
            var rSize = Marshal.SizeOf(rInfo);
            var rBuffer = Marshal.AllocCoTaskMem(rSize);

            Marshal.StructureToPtr(rInfo, rBuffer, true);
            NativeMethods.WinINet.InternetSetOption(IntPtr.Zero, 38, rBuffer, rSize);

            if (rBuffer != IntPtr.Zero)
                Marshal.FreeCoTaskMem(rBuffer);
        }

        public void ClearCache(bool rpClearCookie)
        {
            const int ERROR_NO_MORE_ITEMS = 259;

            var rGroupID = 0L;
            var rHandle = NativeMethods.WinINet.FindFirstUrlCacheGroup(0, 0, IntPtr.Zero, 0, ref rGroupID, IntPtr.Zero);
            if (rHandle != IntPtr.Zero && Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                return;

            var rBufferInitialSize = 0;
            rHandle = NativeMethods.WinINet.FindFirstUrlCacheEntryW(null, IntPtr.Zero, ref rBufferInitialSize);
            if (rHandle != IntPtr.Zero && Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                return;

            var rBufferSize = rBufferInitialSize;
            var rBuffer = Marshal.AllocHGlobal(rBufferSize);
            rHandle = NativeMethods.WinINet.FindFirstUrlCacheEntryW(null, rBuffer, ref rBufferSize);

            while (true)
            {
                var rEntryInfo = (NativeStructs.INTERNET_CACHE_ENTRY_INFO)Marshal.PtrToStructure(rBuffer, typeof(NativeStructs.INTERNET_CACHE_ENTRY_INFO));
                rBufferInitialSize = rBufferSize;

                if (rpClearCookie ||
                    (rEntryInfo.CacheEntryType & NativeEnums.CACHEENTRYTYPE.COOKIE_CACHE_ENTRY) != NativeEnums.CACHEENTRYTYPE.COOKIE_CACHE_ENTRY &&
                    (rEntryInfo.CacheEntryType & NativeEnums.CACHEENTRYTYPE.URLHISTORY_CACHE_ENTRY) != NativeEnums.CACHEENTRYTYPE.URLHISTORY_CACHE_ENTRY)
                    NativeMethods.WinINet.DeleteUrlCacheEntryW(rEntryInfo.lpszSourceUrlName);

                var rResult = NativeMethods.WinINet.FindNextUrlCacheEntryW(rHandle, rBuffer, ref rBufferInitialSize);
                if (!rResult && Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                    break;
                if (!rResult && rBufferInitialSize > rBufferSize)
                {
                    rBufferSize = rBufferInitialSize;
                    rBuffer = Marshal.ReAllocHGlobal(rBuffer, (IntPtr)rBufferSize);
                    NativeMethods.WinINet.FindNextUrlCacheEntryW(rHandle, rBuffer, ref rBufferInitialSize);
                }
            }

            Marshal.FreeHGlobal(rBuffer);
        }

        static int GetIEVersion()
        {
            const string rKey = @"HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer";
            var rVersionText = Registry.GetValue(rKey, "svcVersion", null) as string;
            return rVersionText != null ? Version.Parse(rVersionText).Major * 1000 : 8000;
        }
        static void SetRegistry()
        {
            const string rKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
            var rProcessName = Process.GetCurrentProcess().ProcessName + ".exe";
            if (Registry.GetValue(rKey, rProcessName, null) == null)
                if (OS.IsWin10OrLater)
                    Registry.SetValue(rKey, rProcessName, "11000", RegistryValueKind.DWord);
                else
                    Registry.SetValue(rKey, rProcessName, GetIEVersion(), RegistryValueKind.DWord);
        }
    }
}
