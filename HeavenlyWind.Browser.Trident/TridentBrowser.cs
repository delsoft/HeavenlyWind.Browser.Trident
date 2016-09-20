using mshtml;
using Sakuno.Reflection;
using Sakuno.SystemInterop;
using Sakuno.UserInterface;
using SHDocVw;
using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using WebBrowser = System.Windows.Controls.WebBrowser;

namespace Sakuno.KanColle.Amatsukaze.Browser.Trident
{
    class TridentBrowser : Border, IBrowser
    {
        WebBrowser r_Browser;
        FieldInfo r_WebBrowser2FieldInfo;

        public event Action<bool, bool, string> LoadCompleted = delegate { };

        public TridentBrowser()
        {
            Child = r_Browser = new WebBrowser();

            r_WebBrowser2FieldInfo = typeof(WebBrowser).GetField("_axIWebBrowser2", BindingFlags.Instance | BindingFlags.NonPublic);

            r_Browser.Navigated += (s, e) => ConfigureBrowser();
            r_Browser.LoadCompleted += (s, e) =>
            {
                LoadCompleted(r_Browser.CanGoBack, r_Browser.CanGoForward, e.Uri.ToString());

                ExtractFlash();
            };
        }

        public void GoBack() => r_Browser.GoBack();
        public void GoForward() => r_Browser.GoForward();

        public void Navigate(string rpUrl) => r_Browser.Navigate(rpUrl);

        public void Refresh() => r_Browser.Refresh();

        void ConfigureBrowser()
        {
            var rWebBrowser = r_WebBrowser2FieldInfo?.FastGetValue(r_Browser) as SHDocVw.WebBrowser;
            if (rWebBrowser == null)
                return;

            rWebBrowser.Silent = true;
            rWebBrowser.RegisterAsDropTarget = false;
        }

        public void SetZoom(double rpZoom)
        {
            var rWebBrowser = r_WebBrowser2FieldInfo?.FastGetValue(r_Browser) as SHDocVw.WebBrowser;
            if (rWebBrowser == null)
                return;

            rpZoom = DpiUtil.ScaleX + rpZoom - 1.0;

            object rZoom = (int)(rpZoom * 100);
            rWebBrowser.ExecWB(OLECMDID.OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT.OLECMDEXECOPT_DODEFAULT, ref rZoom);
        }

        public void ExtractFlash()
        {
            try
            {
                IHTMLElement rElement;
                var rDocument = r_Browser.Document as HTMLDocument;
                var rUri = r_Browser.Source;

                if (rUri.AbsoluteUri == GameConstants.GamePageUrl)
                    rElement = rDocument.getElementById("game_frame");
                else
                    rElement = rDocument.getElementsByTagName("EMBED").OfType<IHTMLElement>().SingleOrDefault(r => ((string)r.getAttribute("src")).Contains("kcs/mainD2.swf"));

                if (rElement != null)
                {
                    rElement.document.createStyleSheet().cssText = @"body {
    margin: 0;
    overflow: hidden;
}

#game_frame {
    position: fixed;
    left: 50%;
    top: -16px;
    margin-left: -450px;
    z-index: 255;
}";
                }
            }
            catch
            {
            }
        }

        public ScreenshotData TakeScreenshot()
        {
            var rDocument = r_Browser.Document as HTMLDocument;
            if (rDocument == null)
                return null;

            if (rDocument.url.Contains("kcs/mainD2.swf"))
            {
                var rViewObject = rDocument.getElementsByTagName("EMBED").item(index: 0) as NativeInterfaces.IViewObject;
                if (rViewObject == null)
                    return null;

                return TakeScreenshotCore(rViewObject);
            }
            else
            {
                if (rDocument.getElementById("game_frame") == null)
                    return null;

                var rFrames = rDocument.frames;
                for (var i = 0; i < rFrames.length; i++)
                {
                    var rServiceProvider = rFrames.item(i) as NativeInterfaces.IServiceProvider;
                    if (rServiceProvider == null)
                        return null;

                    var rGuidWebBrowserApp = typeof(IWebBrowserApp).GUID;
                    var rGuidWebBrowser2 = typeof(IWebBrowser2).GUID;
                    var rObject = rServiceProvider.QueryService(ref rGuidWebBrowserApp, ref rGuidWebBrowser2);
                    var rWebBrowser = rObject as IWebBrowser2;
                    if (rWebBrowser == null || rWebBrowser.Document == null)
                        return null;

                    var rViewObject = rWebBrowser.Document.getElementById("externalswf") as NativeInterfaces.IViewObject;
                    if (rViewObject == null)
                        continue;

                    return TakeScreenshotCore(rViewObject);
                }
            }

            return null;
        }
        ScreenshotData TakeScreenshotCore(NativeInterfaces.IViewObject rpViewObject)
        {
            const int BitCount = 24;

            var rEmbedElement = rpViewObject as HTMLEmbed;
            var rWidth = rEmbedElement.clientWidth;
            var rHeight = rEmbedElement.clientHeight;

            var rScreenDC = NativeMethods.User32.GetDC(IntPtr.Zero);
            var rHDC = NativeMethods.Gdi32.CreateCompatibleDC(rScreenDC);

            var rInfo = new NativeStructs.BITMAPINFO();
            rInfo.bmiHeader.biSize = Marshal.SizeOf(typeof(NativeStructs.BITMAPINFOHEADER));
            rInfo.bmiHeader.biWidth = rWidth;
            rInfo.bmiHeader.biHeight = rHeight;
            rInfo.bmiHeader.biBitCount = BitCount;
            rInfo.bmiHeader.biPlanes = 1;

            IntPtr rBits;
            var rHBitmap = NativeMethods.Gdi32.CreateDIBSection(rHDC, ref rInfo, 0, out rBits, IntPtr.Zero, 0);
            var rOldObject = NativeMethods.Gdi32.SelectObject(rHDC, rHBitmap);

            var rTargetDevice = new NativeStructs.DVTARGETDEVICE() { tdSize = 0 };
            var rRect = new NativeStructs.RECT(0, 0, rWidth, rHeight);
            var rEmptyRect = default(NativeStructs.RECT);
            rpViewObject.Draw(1, 0, IntPtr.Zero, ref rTargetDevice, IntPtr.Zero, rHDC, ref rRect, ref rEmptyRect, IntPtr.Zero, IntPtr.Zero);

            var rResult = new ScreenshotData(rWidth, rHeight, BitCount);

            var rPixels = new byte[rWidth * rHeight * 3];
            Marshal.Copy(rBits, rPixels, 0, rPixels.Length);
            rResult.BitmapData = rPixels;

            NativeMethods.Gdi32.SelectObject(rHDC, rOldObject);
            NativeMethods.Gdi32.DeleteObject(rHBitmap);
            NativeMethods.Gdi32.DeleteDC(rHDC);
            NativeMethods.User32.ReleaseDC(IntPtr.Zero, rScreenDC);

            return rResult;
        }
    }
}
