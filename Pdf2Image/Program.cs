using System.Drawing;
using System;
using System.Drawing.Imaging;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Runtime.InteropServices;
using System.Linq;

internal static class PdfiumNative
{
    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern void FPDF_InitLibrary();

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern void FPDF_DestroyLibrary();

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern IntPtr FPDF_LoadDocument(byte[] filePath, string password);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern int FPDF_GetPageCount(IntPtr document);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double FPDF_GetPageWidth(IntPtr page);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern double FPDF_GetPageHeight(IntPtr page);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern void FPDF_ClosePage(IntPtr page);

    [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
    public static extern void FPDF_CloseDocument(IntPtr document);
}

namespace pdf2png
{
    internal static class NativeMethods
    {
        private const string DllName = "pdfium.dll";

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBitmap_CreateEx(int width, int height, int format, IntPtr buffer, int stride);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFBitmap_FillRect(IntPtr bitmap, int left, int top, int width, int height, uint color);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBitmap_GetBuffer(IntPtr bitmap);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFBitmap_Destroy(IntPtr bitmap);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_RenderPageBitmap(IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);
    }

    internal class Program
    {
        public static void WriteErrorLog(Exception ex, string ext)
        {
            try
            {
                string error;
                string stLogDir = GetAppRoot(out error);
                string logFilePath = Path.Combine(stLogDir, "application_log.txt");

                Directory.CreateDirectory(Path.GetDirectoryName(logFilePath)); // 경로 없으면 생성

                using (StreamWriter sw = new StreamWriter(logFilePath, true)) // append mode
                {
                    sw.WriteLine("==== 로그메시지 ====");
                    sw.WriteLine($"시간     : {DateTime.Now}");
                    sw.WriteLine($"메시지   : {ex.Message}");
                    sw.WriteLine($"스택트레이스: {ex.StackTrace}");
                    sw.WriteLine($"추가메시지: {ext}");
                    sw.WriteLine();
                }
            }
            catch
            {
                // 로깅 자체가 실패한 경우는 무시
            }
        }

        public static string GetAppRoot(out string error)
        {
            error = null;
            try
            {
                return Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            }
            catch (IOException e)
            {
                error = "Can't get app root directory\n" + e.StackTrace;
            }
            return null;
        }

        public static string GetProgramFilePath(string relative_path)
        {
            string error;
            string app_root = GetAppRoot(out error);
            return app_root + "\\" + relative_path;
        }

        private static void SaveJpeg(Bitmap bmp, string path, int jpgQuality)
        {
            ImageCodecInfo jpgEncoder = ImageCodecInfo.GetImageDecoders()
                .FirstOrDefault(c => c.FormatID == ImageFormat.Jpeg.Guid);

            if (jpgEncoder != null)
            {
                EncoderParameters encParams = new EncoderParameters(1);
                encParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, jpgQuality);
                bmp.Save(path, jpgEncoder, encParams);
            }
            else
            {
                // fallback
                bmp.Save(path, ImageFormat.Jpeg);
            }
        }

        private static bool HasTransparentPixels(Bitmap bmp)
        {
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            BitmapData data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                unsafe
                {
                    byte* ptr = (byte*)data.Scan0;
                    int stride = data.Stride;

                    for (int y = 0; y < bmp.Height; y++)
                    {
                        for (int x = 0; x < bmp.Width; x++)
                        {
                            byte alpha = ptr[y * stride + x * 4 + 3];
                            if (alpha < 255)
                                return true;
                        }
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(data);
            }

            return false;
        }

        private static void ConvertToWhiteBackground(Bitmap src, Bitmap dest, int alphaThreshold)
        {
            Rectangle rect = new Rectangle(0, 0, src.Width, src.Height);
            BitmapData srcData = src.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData destData = dest.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

            unsafe
            {
                byte* srcPtr = (byte*)srcData.Scan0;
                byte* destPtr = (byte*)destData.Scan0;

                int srcStride = srcData.Stride;
                int destStride = destData.Stride;

                for (int y = 0; y < src.Height; y++)
                {
                    byte* srcRow = srcPtr + y * srcStride;
                    byte* destRow = destPtr + y * destStride;

                    for (int x = 0; x < src.Width; x++)
                    {
                        byte b = srcRow[x * 4 + 0];
                        byte g = srcRow[x * 4 + 1];
                        byte r = srcRow[x * 4 + 2];
                        byte a = srcRow[x * 4 + 3];

                        if (a < alphaThreshold)
                        {
                            destRow[x * 3 + 0] = 255;
                            destRow[x * 3 + 1] = 255;
                            destRow[x * 3 + 2] = 255;
                        }
                        else
                        {
                            float alpha = a / 255f;
                            destRow[x * 3 + 0] = (byte)(b * alpha + 255 * (1 - alpha));
                            destRow[x * 3 + 1] = (byte)(g * alpha + 255 * (1 - alpha));
                            destRow[x * 3 + 2] = (byte)(r * alpha + 255 * (1 - alpha));
                        }
                    }
                }
            }

            src.UnlockBits(srcData);
            dest.UnlockBits(destData);
        }

        static void Main(string[] args)
        {
            String stCurrentDir = System.Environment.CurrentDirectory;

            string appPath = GetProgramFilePath("");

            string arch = IntPtr.Size == 8 ? "x64" : "x86";
            string dllPath = Path.Combine(appPath, arch);
            Environment.SetEnvironmentVariable("PATH", dllPath + ";" + Environment.GetEnvironmentVariable("PATH"));

            string png_AppCode = "002";  // 001 : UBIFORM, 002: MySuit
            string pdf_filename = null;
            string png_filename = null;
           
            float ImgResolutionLevel = pdf2png.Properties.Settings.Default.ImgResolutionLevel;
            float ImgQuality = pdf2png.Properties.Settings.Default.ImgQuality;

            if (args.Length == 2)
            {
                pdf_filename = args[0];
                png_filename = args[1];
            }
            else if (args.Length == 3)
            {
                pdf_filename = args[0];
                png_filename = args[1];
                png_AppCode = args[2];
            }
            else if (args.Length == 4)
            {
                pdf_filename = args[0];
                png_filename = args[1];
                png_AppCode = args[2];
                ImgResolutionLevel = float.Parse(args[3], CultureInfo.InvariantCulture.NumberFormat);
            }
            else if (args.Length == 0)
            {
                pdf_filename = GetProgramFilePath("test.pdf");
                png_filename = GetProgramFilePath("converted.jpg");
            }
            else
            {
                Console.WriteLine("USAGE : pdf2png pdf_filename img_filename_%d");
                Environment.Exit(0);
            }

            String appName = png_AppCode.Equals("001") ? "UBIFORM Editor.exe" : "MySuit Editor.exe";

            String ubiformPath = stCurrentDir.LastIndexOf("UBIReport4Inst") != -1 ?
                stCurrentDir + "\\" + appName : stCurrentDir + "\\UBIReport4Inst\\" + appName;

            // PDF FIle로 부터 페이지 크기 정보를 얻어온다.
            System.IO.FileInfo inputPdf = new FileInfo(pdf_filename);
            try
            {
                PdfiumNative.FPDF_InitLibrary();
            }
            catch(Exception ex)
            {
                WriteErrorLog(ex, dllPath);
                System.Diagnostics.Process.Start(ubiformPath, "pdf2image FAIL");
                Environment.Exit(0);
            }

            byte[] utf8Path = System.Text.Encoding.UTF8.GetBytes(inputPdf.FullName + "\0");
            IntPtr doc = PdfiumNative.FPDF_LoadDocument(utf8Path, null);
            if (doc == IntPtr.Zero)
            {
                WriteErrorLog(new Exception("PDF 문서를 열 수 없습니다."), "");
                System.Diagnostics.Process.Start(ubiformPath, "pdf2image FAIL");
                Environment.Exit(0);
            }

            int PrintQuality = ImgQuality > 0 && ImgQuality <= 1.0 ? (int)Math.Round(ImgQuality * 100) : 80;

            StringBuilder stbPageWidthInfo = new StringBuilder();
            StringBuilder stbPageHeightInfo = new StringBuilder();

            int numberOfPages = PdfiumNative.FPDF_GetPageCount(doc);

            for (int i = 0; i < numberOfPages; i++)
            {
                IntPtr page = PdfiumNative.FPDF_LoadPage(doc, i);
                double pageWidth = PdfiumNative.FPDF_GetPageWidth(page);
                double pageHeight = PdfiumNative.FPDF_GetPageHeight(page);

                int dpi = (int) Math.Round(96 * ImgResolutionLevel);
                int width = (int)(pageWidth * dpi / 72.0);
                int height = (int)(pageHeight * dpi / 72.0);
                Console.WriteLine($"Rendering page {i + 1}: {width}x{height}");

                if (i == 0)
                {
                    stbPageWidthInfo.Append(String.Format("{0}", (int)(pageWidth * (96f / 72f))));
                    stbPageHeightInfo.Append(String.Format("{0}", (int)(pageHeight * (96f / 72f))));
                }
                else if (i > 0)
                {
                    stbPageWidthInfo.Append(String.Format(",{0}", (int)(pageWidth * (96f / 72f))));
                    stbPageHeightInfo.Append(String.Format(",{0}", (int)(pageHeight * (96f / 72f))));
                }

                IntPtr bitmap = NativeMethods.FPDFBitmap_CreateEx(width, height, 4, IntPtr.Zero, width * 4);
                NativeMethods.FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xFFFFFFFF);

                int flags = 0x01 | 0x02 | 0x10; // FPDF_ANNOT | FPDF_LCD_TEXT | FPDF_NO_CATCH
                NativeMethods.FPDF_RenderPageBitmap(bitmap, page, 0, 0, width, height, 0, flags);

                IntPtr buffer = NativeMethods.FPDFBitmap_GetBuffer(bitmap);
                int length = width * height * 4;
                byte[] pixelData = new byte[length];
                Marshal.Copy(buffer, pixelData, 0, length);

                for (int p = 0; p < length; p += 4)
                {
                    byte b = pixelData[p + 0];
                    byte g = pixelData[p + 1];
                    byte r = pixelData[p + 2];
                    byte a = pixelData[p + 3];

                    pixelData[p + 0] = r;
                    pixelData[p + 1] = g;
                    pixelData[p + 2] = b;
                    pixelData[p + 3] = a;
                }

                String newPng_filename = png_filename.Substring(0, png_filename.LastIndexOf("."));

                using (Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb))
                {
                    Rectangle rect = new Rectangle(0, 0, width, height);
                    BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    Marshal.Copy(pixelData, 0, bmpData.Scan0, length);
                    bmp.UnlockBits(bmpData);

                    //string outputPath = Path.Combine(newPng_filename, $"_{i + 1}.jpg");
                    string outputPath = newPng_filename + "_" + (i + 1) + ".jpg";

                    if (HasTransparentPixels(bmp))
                    {
                        using (Bitmap finalBmp = new Bitmap(width, height, PixelFormat.Format24bppRgb))
                        {
                            ConvertToWhiteBackground(bmp, finalBmp, 32);
                            SaveJpeg(finalBmp, outputPath, PrintQuality);
                            Console.WriteLine($"[투명 → 흰배경] 저장 완료: {outputPath}");
                        }
                    }
                    else
                    {
                        SaveJpeg(bmp, outputPath, PrintQuality);
                        Console.WriteLine($"[일반 배경] 저장 완료: {outputPath}");
                    }
                }

                NativeMethods.FPDFBitmap_Destroy(bitmap);
                PdfiumNative.FPDF_ClosePage(page);
            }

            PdfiumNative.FPDF_CloseDocument(doc);
            PdfiumNative.FPDF_DestroyLibrary();

            Console.WriteLine("Conversion is successful.");
            WriteErrorLog(new Exception("Conversion is successful."), "");

            try
            {
                String sResult = String.Format("pdf2image SUCCESS {0} {1}", stbPageWidthInfo.ToString(), stbPageHeightInfo.ToString());
                System.Diagnostics.Process.Start(ubiformPath, sResult);
            }
            catch (Exception ex)
            {
                WriteErrorLog(ex, "");
                Environment.Exit(0);
            }
        }
    }
}
