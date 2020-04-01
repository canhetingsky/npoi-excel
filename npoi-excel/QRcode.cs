using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ZXing;

namespace npoi_excel
{
    class QRcode
    {
        public static byte[] Encode(string msg,int codeSizeInPixels = 100)
        {
            BarcodeWriter writer = new BarcodeWriter();
            writer.Format = BarcodeFormat.QR_CODE;
            writer.Options.Hints.Add(EncodeHintType.CHARACTER_SET, "UTF-8");//编码问题
            writer.Options.Hints.Add(
                EncodeHintType.ERROR_CORRECTION,
                ZXing.QrCode.Internal.ErrorCorrectionLevel.H
            );

            writer.Options.Height = writer.Options.Width = codeSizeInPixels;    //设置图片长宽
            writer.Options.Margin = 1;//设置边框
            ZXing.Common.BitMatrix bm = writer.Encode(msg);
            Bitmap bitmap = writer.Write(bm);

            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                byte[] bytes = new byte[stream.Length];
                stream.Seek(0, SeekOrigin.Begin);
                stream.Read(bytes, 0, Convert.ToInt32(stream.Length));
                return bytes;
            }
        }
    }
}
